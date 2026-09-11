const { day, cutoff } = require('./analytics-store');
const bad = message => Object.assign(new Error(message), { status: 400 });
const guid = value => typeof value === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
function parseAdmins(text = '[]') {
  const entries = JSON.parse(text);
  if (!Array.isArray(entries) || entries.some(e => !e || !guid(e.tenant_id) || !(e.user_oid !== undefined
    ? guid(e.user_oid) && e.email === undefined : typeof e.email === 'string' && /^[^\s@]+@[^\s@]+$/.test(e.email))
    || !Array.isArray(e.tenants) || !e.tenants.length || e.tenants.some(t => t !== '*' && !guid(t)))) throw new Error('Invalid ANALYTICS_ADMINS');
  return entries.map(e => ({ tenant_id: e.tenant_id.toLowerCase(), ...(e.user_oid ? { user_oid:e.user_oid.toLowerCase() } : { email:e.email.toLowerCase() }), tenants: e.tenants.map(t => t.toLowerCase()) }));
}
function allowedTenants(identity, admins = []) {
  return [...new Set(admins.filter(e => e.tenant_id === identity.tid.toLowerCase() && e.user_oid === identity.oid.toLowerCase()).flatMap(e => e.tenants))];
}
function filters(query, tenants, now = Date.now()) {
  const string = key => { const value = query[key]; if (value !== undefined && typeof value !== 'string') throw bad('筛选参数格式无效。'); return value || ''; };
  const to = string('to') || day(now);
  const from = string('from') || day(now - 29 * 86400000);
  const dateValid = value => /^\d{4}-\d{2}-\d{2}$/.test(value) && Number.isFinite(Date.parse(`${value}T00:00:00Z`)) && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
  if (!dateValid(from) || !dateValid(to) || from > to || from < cutoff(now) || to > day(now)) throw bad('请选择保留期内的日期，截止日期不能晚于今天。');
  const tenant = string('tenant').toLowerCase(), host = string('host'), clientType = string('client_type'), userOid = string('user_oid'), search = string('search').trim();
  if (tenant && !tenants.includes('*') && !tenants.includes(tenant)) throw Object.assign(new Error('无权查询此租户。'), { status: 403 });
  if (tenant && !guid(tenant)) throw bad('租户 ID 格式无效。');
  if (host && !['word', 'outlook', 'unknown'].includes(host)) throw bad('平台无效。');
  if (clientType && !['online', 'local', 'unknown'].includes(clientType)) throw bad('客户端类型无效。');
  if (userOid.length > 100 || search.length > 100) throw bad('搜索内容过长。');
  const page = Number(string('page') || 1), pageSize = Number(string('page_size') || 50);
  if (!Number.isSafeInteger(page) || page < 1 || page > 1000000 || !Number.isSafeInteger(pageSize) || pageSize < 1 || pageSize > 200) throw bad('分页参数无效。');
  return { from, to, tenants: tenant ? [tenant] : tenants, host, clientType, userOid, search, page, pageSize };
}
function registerAdmin(app, config, analytics, services) {
  const profiles = new Map();
  app.use('/api/admin', async (req, res, next) => {
    req.adminTenants = allowedTenants(req.identity, config.analyticsAdmins);
    const candidates = (config.analyticsAdmins || []).filter(e => e.email && e.tenant_id === req.identity.tid.toLowerCase());
    if (!req.adminTenants.length && candidates.length) {
      try {
        const key = `${req.identity.tid}:${req.identity.oid}`;
        let cached = profiles.get(key);
        if (!cached || cached.until <= Date.now()) {
          const profile = await services.profile(req.token,req.identity);
          if (profile.id !== req.identity.oid || profile.tenantId !== req.identity.tid) throw new Error('Identity mismatch');
          cached = { mail: typeof profile.mail === 'string' ? profile.mail.toLowerCase() : '', until:Date.now()+300000 };
          if (profiles.size >= 100) profiles.delete(profiles.keys().next().value);
          profiles.set(key,cached);
        }
        req.adminTenants = [...new Set(candidates.filter(e => e.email === cached.mail).flatMap(e => e.tenants))];
      } catch { return res.status(403).json({error:'无法核验管理员邮箱。请确认此账号的用户资料权限及同意，或由维护人员配置真实对象 ID。'}); }
    }
    if (!req.adminTenants.length) return res.status(403).json({ error: '此账号没有统计后台查询权限。' });
    next();
  });
  app.get('/api/admin/session', (req, res) => res.json({ tenants: req.adminTenants }));
  for (const [route, mode] of [['overview', 'usage'], ['api-usage', 'api']]) {
    app.get(`/api/admin/${route}`, async (req, res) => {
      try { res.json(await analytics.query(filters(req.query, req.adminTenants), mode)); }
      catch (error) { res.status(error.status || 503).json({ error: error.status ? error.message : '统计暂不可用，请稍后重试。翻译功能不受影响。' }); }
    });
  }
}
module.exports = { parseAdmins, allowedTenants, filters, registerAdmin };

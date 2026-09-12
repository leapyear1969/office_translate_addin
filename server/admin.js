const { day, cutoff } = require('./analytics-store');
const bad = message => Object.assign(new Error(message), { status: 400 });
const guid = value => typeof value === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
function parseAdmins(text = '[]') {
  const entries = JSON.parse(text);
  if (!Array.isArray(entries) || entries.some(e => !e || !guid(e.tenant_id) || !(e.user_oid !== undefined
    ? guid(e.user_oid) && e.email === undefined : typeof e.email === 'string' && /^[^\s@]+@[^\s@]+$/.test(e.email))
    || (e.can_manage_admins !== undefined && typeof e.can_manage_admins !== 'boolean')
    || !Array.isArray(e.tenants) || !e.tenants.length || e.tenants.some(t => t !== '*' && !guid(t)))) throw new Error('Invalid ANALYTICS_ADMINS');
  return entries.map(e => ({ tenant_id: e.tenant_id.toLowerCase(), ...(e.user_oid ? { user_oid:e.user_oid.toLowerCase() } : { email:e.email.toLowerCase() }), tenants: [...new Set(e.tenants.map(t => t.toLowerCase()))], ...(e.can_manage_admins !== undefined ? {can_manage_admins:e.can_manage_admins} : {}) }));
}
function allowedTenants(identity, admins = []) {
  return [...new Set(admins.filter(e => e.tenant_id === identity.tid.toLowerCase() && e.user_oid === identity.oid.toLowerCase()).flatMap(e => e.tenants))];
}
function filters(query, tenants, now = Date.now(), retentionMonths) {
  const string = key => { const value = query[key]; if (value !== undefined && typeof value !== 'string') throw bad('筛选参数格式无效。'); return value || ''; };
  const to = string('to') || day(now);
  const retainedFrom = cutoff(now, retentionMonths);
  const defaultFrom = day(now - 29 * 86400000);
  const from = string('from') || (retainedFrom !== null && defaultFrom < retainedFrom ? retainedFrom : defaultFrom);
  const dateValid = value => /^\d{4}-\d{2}-\d{2}$/.test(value) && Number.isFinite(Date.parse(`${value}T00:00:00Z`)) && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
  if (!dateValid(from) || !dateValid(to) || from > to || (retainedFrom !== null && from < retainedFrom) || to > day(now)) throw bad('请选择保留期内的日期，截止日期不能晚于今天。');
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
  const seed = (config.analyticsAdmins || []).map(e => ({...e,can_manage_admins:e.can_manage_admins ?? e.tenants.includes('*')}));
  const keyOf = e => `${e.tenant_id}:${e.user_oid ? `oid:${e.user_oid}` : `email:${e.email}`}`;
  app.use('/api/admin', async (req, res, next) => {
    try {
      req.adminSettings = await analytics.adminSettings(seed);
    } catch { return res.status(503).json({error:'管理员配置暂不可用，请稍后重试。'}); }
    const entries = req.adminSettings.entries;
    // Mail/UPN can be reassigned. Legacy email entries remain visible for
    // manual migration but must never grant access to a new directory object.
    const matched = entries.filter(e => e.tenant_id === req.identity.tid.toLowerCase() && e.user_oid === req.identity.oid.toLowerCase());
    req.adminTenants = [...new Set(matched.flatMap(e => e.tenants))];
    req.canManageAdmins = matched.some(e => e.can_manage_admins === true);
    req.adminKeys = matched.map(keyOf);
    if (!req.adminTenants.length) return res.status(403).json({ error: '此账号没有统计后台查询权限。' });
    next();
  });
  app.get('/api/admin/session', (req, res) => res.json({ tenants: req.adminTenants, canManageAdmins:req.canManageAdmins, retained_from:cutoff(Date.now(), config.analyticsRetentionMonths) }));
  app.get('/api/admin/settings', (req,res) => {
    if (!req.canManageAdmins) return res.status(403).json({error:'此账号没有管理员配置权限。'});
    res.json({...req.adminSettings,selfKeys:req.adminKeys});
  });
  app.put('/api/admin/settings', async (req,res) => {
    if (!req.canManageAdmins) return res.status(403).json({error:'此账号没有管理员配置权限。'});
    try {
      if (!Number.isSafeInteger(req.body?.revision) || req.body.revision < 1) throw bad('配置版本无效，请重新加载。');
      if (req.body.revision !== req.adminSettings.revision) throw Object.assign(new Error('配置已被其他管理员更新，请重新加载后再修改。'),{status:409});
      let entries;
      try { entries = parseAdmins(JSON.stringify(req.body.entries)); } catch { throw bad('请输入有效的登录租户 ID、对象 ID，以及可查看租户。'); }
      if (entries.some(e => !e.user_oid)) throw bad('邮箱不能作为授权标识，请核对原获批账号并填写用户对象 ID。');
      if (!entries.length || entries.length > 100) throw bad('请保留 1 至 100 位管理员。');
      if (new Set(entries.map(keyOf)).size !== entries.length) throw bad('同一登录租户中的管理员账号不能重复。');
      if (!entries.some(e => e.can_manage_admins === true)) throw bad('至少保留一位可以管理管理员的账号。');
      if (!entries.some(e => req.adminKeys.includes(keyOf(e)) && e.can_manage_admins === true)) throw bad('不能移除自己的管理员配置权限或修改自己的登录身份。');
      const saved = await analytics.saveAdminSettings({revision:req.body.revision,entries,
        actor:{tenant_id:req.identity.tid.toLowerCase(),user_oid:req.identity.oid.toLowerCase()}});
      res.json({...saved,selfKeys:req.adminKeys});
    } catch (error) { res.status(error.status || 503).json({error:error.status ? error.message : '保存暂不可用，请重新加载确认配置后重试。'}); }
  });
  for (const [route, mode] of [['overview', 'usage'], ['api-usage', 'api']]) {
    app.get(`/api/admin/${route}`, async (req, res) => {
      try { res.json(await analytics.query(filters(req.query, req.adminTenants, Date.now(), config.analyticsRetentionMonths), mode)); }
      catch (error) { res.status(error.status || 503).json({ error: error.status ? error.message : '统计暂不可用，请稍后重试。翻译功能不受影响。' }); }
    });
  }
}
module.exports = { parseAdmins, allowedTenants, filters, registerAdmin };

const express = require('express');
const { rateLimit } = require('express-rate-limit');
const { createAuth } = require('./auth');
const { createTranslator } = require('./translator');
const { createConsent } = require('./consent');
const { createAnalytics, requestContext, recordRequest, context } = require('./analytics');
const { registerAdmin } = require('./admin');
const { adminPageSecurity } = require('./admin-security');

function createApp(config, dependencies) {
  const services = dependencies || { ...createAuth(config), ...createTranslator(config) };
  const app = express();
  const analytics = dependencies?.analytics || createAnalytics(config.analyticsDatabaseUrl, undefined, config.analyticsRetentionMonths);
  app.locals.analytics = analytics;
  const consent = createConsent(config);
  app.disable('x-powered-by');
  app.use(adminPageSecurity(config));
  app.get('/admin/config', (_req, res) => res.set('Cache-Control', 'no-store').json({
    configured: !!(config.adminClientId && config.resource), clientId: config.adminClientId || '',
    authority: `${config.authority}/${config.adminTenant || 'organizations'}`,
    // A shared SPA/API registration uses its GUID for self-resource tokens.
    scope: config.resource ? `${config.adminClientId === config.clientId ? config.clientId : config.resource.replace(/\/$/, '')}/access_as_user` : '',
    redirectUri: `${config.origin}/admin/usage`,
  }));
  app.get(['/admin/usage', '/admin/api', '/admin/settings'], (req, res, next) => {
    req.url = '/admin.html'; next();
  });
  app.use('/auth/consent', (_req, res, next) => {
    res.set({ 'Cache-Control': 'no-store', 'Referrer-Policy': 'no-referrer' }); next();
  }, rateLimit({ windowMs: 60000, limit: 60 }));
  app.get('/auth/consent/launch/:state', (req, res, next) => {
    try { res.json(consent.launch(req.params.state)); } catch (error) { next(error); }
  });
  app.get('/auth/consent/callback', async (req, res, next) => {
    try { res.redirect(await consent.complete(req.query)); } catch (error) { next(error); }
  });
  app.use('/api', (req, res, next) => {
    res.set('Cache-Control', 'no-store');
    if (req.get('Origin') && req.get('Origin') !== config.origin) return res.status(403).json({ error: '请求来源不被允许。' });
    next();
  });
  const anonymousAdminLimit = rateLimit({ windowMs: 60000, limit: 120, standardHeaders: 'draft-7', legacyHeaders: false,
    message: { error: '未认证请求过于频繁，请稍后重试。' } });
  // Anonymous traffic cannot spend a verified user's quota, even behind a shared proxy.
  app.use('/api/admin', (req, res, next) => /^Bearer (\S+)$/.test(req.get('Authorization') || '')
    ? next() : anonymousAdminLimit(req, res, next));
  app.use('/api', rateLimit({ windowMs: 60000, limit: 60, standardHeaders: 'draft-7', legacyHeaders: false, skip: req => /^\/admin(?:\/|$)/.test(req.path),
    message: { error: '请求过于频繁，请稍后重试。' } }));
  app.get('/health', (_req, res) => res.json({ ok: true }));
  app.use('/api', async (req, res, next) => {
    const match = /^Bearer (\S+)$/.exec(req.get('Authorization') || '');
    if (!match) return res.status(401).json({ error: '请先通过 Office SSO 登录。' });
    try { req.token = match[1]; req.identity = await services.authenticate(req.token); next(); }
    catch (error) {
      const reject = () => res.status(error.status || 401).json({ error: error.status ? error.message : 'SSO 认证失败，请重新登录。' });
      if (/^\/admin(?:\/|$)/.test(req.path)) return anonymousAdminLimit(req, res, reject);
      reject();
    }
  });
  app.use('/api/admin', rateLimit({ windowMs: 60000, limit: 120, standardHeaders: 'draft-7', legacyHeaders: false,
    keyGenerator: req => JSON.stringify([req.identity.tid.toLowerCase(), req.identity.oid.toLowerCase()]),
    message: { error: '请求过于频繁，请稍后重试。' } }));
  app.use('/api', express.json({ limit: '6mb' }));
  registerAdmin(app, config, analytics, services);
  const route = fn => async (req, res, next) => { try { await fn(req, res); } catch (e) { next(e); } };
  app.post('/api/consent/start', route(async (req, res) => res.json(await consent.start(req.identity))));
  app.get('/api/me', route(async (req, res) => {
    const profile = await services.profile(req.token, req.identity);
    try { analytics.profile({ tenantId: req.identity.tid.toLowerCase(), userOid: req.identity.oid.toLowerCase(),
      displayName: typeof profile.displayName === 'string' ? profile.displayName.slice(0, 256) : null,
      organizationName: typeof profile.organizationName === 'string' ? profile.organizationName.slice(0, 256) : null,
      mail: typeof profile.mail === 'string' ? profile.mail.slice(0, 320) : null }); } catch { /* Statistics are optional to login. */ }
    res.json(profile);
  }));
  async function business(req, endpoint, run) {
    return context.run(requestContext(req.identity, req.body, analytics), async () => {
      const started = Date.now(); let success = false;
      try { const result = await run(); success = true; return result; }
      finally { recordRequest('business', endpoint, started, success); }
    });
  }
  function validateHtml(req, res) {
    if (typeof req.body?.html !== 'string' || !req.body.html || req.body.html.length > 1000000) {
      res.status(400).json({ error: '翻译内容为空或超过 1,000,000 字符限制。' }); return false;
    }
    return true;
  }
  app.post('/api/translate', route(async (req, res) => {
    if (!validateHtml(req, res)) return;
    if (typeof req.body.to !== 'string' || !/^[a-z]{2,3}(?:-[A-Za-z]{2,8})?$/.test(req.body.to)) {
      return res.status(400).json({ error: '目标语言无效。' });
    }
    const { subject, html, to } = req.body;
    if (subject !== undefined && (typeof subject !== 'string' || subject.length > 10000)) {
      return res.status(400).json({ error: '邮件标题无效或过长。' });
    }
    res.json(await business(req, 'translate', async () => {
      const translated = { html: await services.translate(html, to) };
      if (subject !== undefined) translated.subject = await services.translateSubject(subject, to);
      return translated;
    }));
  }));
  app.post('/api/translate/word', route(async (req, res) => {
    const { paragraphs, to } = req.body || {};
    if (!Array.isArray(paragraphs) || !paragraphs.length || paragraphs.length > 10000
      || paragraphs.some(value => typeof value !== 'string' || !value.trim() || value.length > 40000)
      || paragraphs.join('').length > 1000000
      || typeof to !== 'string' || !/^[a-z]{2,3}(?:-[A-Za-z]{2,8})?$/.test(to)) {
      return res.status(400).json({ error: '正文段落或目标语言无效，或内容超过限制。' });
    }
    res.json({ paragraphs: await business(req, 'translate-word', () => services.translateWord(paragraphs, to)) });
  }));
  app.post('/api/detect', route(async (req, res) => {
    if (validateHtml(req, res)) res.json(await business(req, 'detect', () => services.detect(req.body.html)));
  }));
  app.use('/api', (_req, res) => res.status(404).json({ error: '接口不存在。' }));
  app.use((error, _req, res, _next) => {
    const status = error.type === 'entity.too.large' ? 413 : error instanceof SyntaxError ? 400 : error.status || 502;
    res.status(status).json({ code: error.code === 'consent_required' ? error.code : undefined, error: status === 413 ? '翻译内容过大。' : status === 400 ? '请求格式无效。'
      : error.status ? error.message : error.name === 'TimeoutError' ? '请求超时，请重试。' : '翻译服务请求失败，请稍后重试。' });
  });
  return app;
}
module.exports = { createApp };

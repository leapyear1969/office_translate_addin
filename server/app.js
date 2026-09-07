const express = require('express');
const { rateLimit } = require('express-rate-limit');
const { createAuth } = require('./auth');
const { createTranslator } = require('./translator');
const { createConsent } = require('./consent');

function createApp(config, dependencies) {
  const services = dependencies || { ...createAuth(config), ...createTranslator(config) };
  const app = express();
  const consent = createConsent(config);
  app.disable('x-powered-by');
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
  app.use('/api', rateLimit({ windowMs: 60000, limit: 60, standardHeaders: 'draft-7', legacyHeaders: false,
    message: { error: '请求过于频繁，请稍后重试。' } }));
  app.get('/health', (_req, res) => res.json({ ok: true }));
  app.use('/api', async (req, res, next) => {
    const match = /^Bearer (\S+)$/.exec(req.get('Authorization') || '');
    if (!match) return res.status(401).json({ error: '请先通过 Outlook SSO 登录。' });
    try { req.token = match[1]; req.identity = await services.authenticate(req.token); next(); }
    catch (error) { res.status(error.status || 401).json({ error: error.status ? error.message : 'SSO 认证失败，请重新登录。' }); }
  });
  app.use('/api', express.json({ limit: '6mb' }));
  const route = fn => async (req, res, next) => { try { await fn(req, res); } catch (e) { next(e); } };
  app.post('/api/consent/start', route(async (req, res) => res.json(await consent.start(req.identity))));
  app.get('/api/me', route(async (req, res) => res.json(await services.profile(req.token, req.identity))));
  function validateHtml(req, res) {
    if (typeof req.body?.html !== 'string' || !req.body.html || req.body.html.length > 1000000) {
      res.status(400).json({ error: '邮件正文为空或超过 1,000,000 字符限制。' }); return false;
    }
    return true;
  }
  app.post('/api/translate', route(async (req, res) => {
    if (!validateHtml(req, res)) return;
    if (typeof req.body.to !== 'string' || !/^[a-z]{2,3}(?:-[A-Za-z]{2,8})?$/.test(req.body.to)) {
      return res.status(400).json({ error: '目标语言无效。' });
    }
    res.json({ html: await services.translate(req.body.html, req.body.to) });
  }));
  app.post('/api/detect', route(async (req, res) => {
    if (validateHtml(req, res)) res.json(await services.detect(req.body.html));
  }));
  app.use('/api', (_req, res) => res.status(404).json({ error: '接口不存在。' }));
  app.use((error, _req, res, _next) => {
    const status = error.type === 'entity.too.large' ? 413 : error instanceof SyntaxError ? 400 : error.status || 502;
    res.status(status).json({ code: error.code === 'consent_required' ? error.code : undefined, error: status === 413 ? '邮件正文过大。' : status === 400 ? '请求格式无效。'
      : error.status ? error.message : error.name === 'TimeoutError' ? '请求超时，请重试。' : '翻译服务请求失败，请稍后重试。' });
  });
  return app;
}
module.exports = { createApp };

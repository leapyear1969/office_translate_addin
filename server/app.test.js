const { createApp } = require('./app');
const config = { origin: 'https://localhost:3000' };
async function setup(run) {
  const translate = jest.fn(async () => '<p>译文</p>');
  const authenticate = jest.fn(async token => { if (token !== 'valid') throw new Error('invalid'); return { oid: 'user', tid: 'tenant' }; });
  const profile = jest.fn(async () => ({ id: 'user', displayName: 'Test', mail: 'test@example.com' }));
  const app = createApp(config, { authenticate, profile, translate, detect: async () => ({ language: 'en', score: 1 }) });
  const server = await new Promise(resolve => { const s = app.listen(0, '127.0.0.1', () => resolve(s)); });
  const request = (path, { body, token, origin } = {}) => fetch(`http://127.0.0.1:${server.address().port}${path}`, {
    method: body ? 'POST' : 'GET', headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(origin ? { Origin: origin } : {}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  try { await run({ request, translate, profile }); }
  finally { server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
}
test('requires SSO before submitting content to Translator', () => setup(async ({ request, translate }) => {
  const r = await request('/api/translate', { body: { html: '<p>Hello</p>', to: 'zh-Hans' } });
  expect(r.status).toBe(401);
  expect(translate).not.toHaveBeenCalled();
}));
test('translates authenticated requests and returns the Graph profile', () => setup(async ({ request }) => {
  const r = await request('/api/translate', { token: 'valid', body: { html: '<p>Hello</p>', to: 'zh-Hans' } });
  expect(r.status).toBe(200);
  expect((await r.json()).html).toBe('<p>译文</p>');
  const me = await request('/api/me', { token: 'valid' });
  expect(me.status).toBe(200);
  expect((await me.json()).displayName).toBe('Test');
}));
test('rejects invalid input and foreign origins before translation', () => setup(async ({ request, translate }) => {
  const invalid = await request('/api/translate', { token: 'valid', body: { html: '<p>x</p>', to: '../bad' } });
  expect(invalid.status).toBe(400);
  const foreign = await request('/api/translate', { origin: 'https://other.example', token: 'valid', body: { html: '<p>x</p>', to: 'en' } });
  expect(foreign.status).toBe(403);
  expect(translate).not.toHaveBeenCalled();
}));

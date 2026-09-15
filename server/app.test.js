const { createApp } = require('./app');
const config = { origin: 'https://localhost:3000' };
test('returns the structured quota error to Office clients', () => setup(async ({ request, translate }) => {
  translate.mockRejectedValue(Object.assign(new Error('本月云翻译额度已用完'), { status: 403, code: 'cloud_quota_exhausted' }));
  const response = await request('/api/translate', { token: 'valid', body: { html: '<p>Hello</p>', to: 'en' } });
  expect(response.status).toBe(403);
  expect(await response.json()).toEqual({ code: 'cloud_quota_exhausted', error: '本月云翻译额度已用完' });
}));
async function setup(run) {
  const translate = jest.fn(async () => '<p>译文</p>');
  const translateSubject = jest.fn(async () => '会议通知');
  const translateWord = jest.fn(async paragraphs => paragraphs.map(p => p.replace('Hello', '你好')));
  const authenticate = jest.fn(async token => { if (token !== 'valid') throw new Error('invalid'); return { oid: 'user', tid: 'tenant' }; });
  const profile = jest.fn(async () => ({ id: 'user', displayName: 'Test', mail: 'test@example.com' }));
  const app = createApp(config, { authenticate, profile, translate, translateSubject, translateWord, detect: async () => ({ language: 'en', score: 1 }) });
  const server = await new Promise(resolve => { const s = app.listen(0, '127.0.0.1', () => resolve(s)); });
  const request = (path, { body, token, origin } = {}) => fetch(`http://127.0.0.1:${server.address().port}${path}`, {
    method: body ? 'POST' : 'GET', headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(origin ? { Origin: origin } : {}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  try { await run({ request, translate, translateSubject, translateWord, profile }); }
  finally { server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
}

test('translates the subject separately as text with the body', () => setup(async ({ request, translateSubject }) => {
  const response = await request('/api/translate', { token: 'valid', body: { html: '<p>Hello</p>', subject: 'Meeting <notice>', to: 'zh-Hans' } });
  expect(response.status).toBe(200);
  expect(await response.json()).toEqual({ html: '<p>译文</p>', subject: '会议通知' });
  expect(translateSubject).toHaveBeenCalledWith('Meeting <notice>', 'zh-Hans');
}));

test.each([null, 12, {}, 'x'.repeat(10001)])('rejects invalid subject before translation (%#)', subject => setup(async ({ request, translate, translateSubject }) => {
  const response = await request('/api/translate', { token: 'valid', body: { html: '<p>Hello</p>', subject, to: 'en' } });
  expect(response.status).toBe(400);
  expect(translate).not.toHaveBeenCalled();
  expect(translateSubject).not.toHaveBeenCalled();
}));
test('requires SSO before submitting content to Translator', () => setup(async ({ request, translate }) => {
  const r = await request('/api/translate', { body: { html: '<p>Hello</p>', to: 'zh-Hans' } });
  expect(r.status).toBe(401);
  expect(translate).not.toHaveBeenCalled();
}));

test('Word paragraph endpoint preserves markers and requires authentication', () => setup(async ({ request, translateWord, translate }) => {
  const body = { paragraphs: ['<p><span id="r0">Hello</span></p>'], to: 'zh-Hans' };
  expect((await request('/api/translate/word', { body })).status).toBe(401);
  expect(translateWord).not.toHaveBeenCalled();
  const response = await request('/api/translate/word', { body, token: 'valid' });
  expect(response.status).toBe(200);
  expect(await response.json()).toEqual({ paragraphs: ['<p><span id="r0">你好</span></p>'] });
  expect(translateWord).toHaveBeenCalledWith(body.paragraphs, body.to);
  expect(translate).not.toHaveBeenCalled();
}));

test.each([null, [], [null], [''], ['x'.repeat(40001)], Array(26).fill('x'.repeat(40000))])('rejects invalid Word paragraph input (%#)', paragraphs => setup(async ({ request, translateWord }) => {
  const response = await request('/api/translate/word', { token: 'valid', body: { paragraphs, to: 'en' } });
  expect(response.status).toBe(400);
  expect(translateWord).not.toHaveBeenCalled();
}));
test('translates authenticated requests and returns the Graph profile', () => setup(async ({ request }) => {
  const r = await request('/api/translate', { token: 'valid', body: { html: '<p>Hello</p>', to: 'zh-Hans' } });
  expect(r.status).toBe(200);
  expect((await r.json()).html).toBe('<p>译文</p>');
  const me = await request('/api/me', { token: 'valid' });
  expect(me.status).toBe(200);
  expect((await me.json()).displayName).toBe('Test');
}));

test('returns a structured consent code and protects the consent start endpoint', () => setup(async ({ request, profile }) => {
  profile.mockRejectedValue(Object.assign(new Error('请授权'), { status: 403, code: 'consent_required' }));
  const me = await request('/api/me', { token: 'valid' });
  expect(me.status).toBe(403);
  expect(await me.json()).toEqual({ error: '请授权', code: 'consent_required' });
  const start = await request('/api/consent/start', { body: {} });
  expect(start.status).toBe(401);
}));
test('rejects invalid input and foreign origins before translation', () => setup(async ({ request, translate }) => {
  const invalid = await request('/api/translate', { token: 'valid', body: { html: '<p>x</p>', to: '../bad' } });
  expect(invalid.status).toBe(400);
  const foreign = await request('/api/translate', { origin: 'https://other.example', token: 'valid', body: { html: '<p>x</p>', to: 'en' } });
  expect(foreign.status).toBe(403);
  expect(translate).not.toHaveBeenCalled();
}));

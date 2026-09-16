const { createTranslator } = require('./translator');
const { readConfig } = require('./config');
const { context } = require('./analytics');

const originalFetch = global.fetch;
const english = 'Hello, how are you?';
const routes = [{ code: 'en', targets: ['zh'] }, { code: 'zh', targets: ['en'] }];
const config = { localTranslatorEndpoint: 'http://argos.test:30261', translatorEndpoint: 'https://azure.test/', translatorKey: 'key' };
const json = body => ({ ok: true, json: async () => body });
function mockLocal() {
  global.fetch = jest.fn(async (url, options) => {
    if (url.pathname === '/languages') return json(routes);
    const body = JSON.parse(options.body);
    return json({ translated_text: body.source === 'en' ? '你好，你怎么样？ <安全>' : english, source: body.source, target: body.target });
  });
}
afterEach(() => { global.fetch = originalFetch; });

test('configuration prefers LAN service by default and supports Azure-only and local-only modes', () => {
  expect(readConfig({}).localTranslatorEndpoint).toBe('http://192.168.3.101:30261');
  expect(readConfig({ LOCAL_TRANSLATOR_ENDPOINT: '' }).localTranslatorEndpoint).toBe('');
  expect(readConfig({ LOCAL_TRANSLATOR_FALLBACK: 'false' }).localTranslatorFallback).toBe(false);
});

test('local subject translation needs no Azure key, maps Simplified Chinese and caches language routes', async () => {
  mockLocal();
  const service = createTranslator({ ...config, translatorKey: '' });
  expect(await service.translateSubject(english, 'zh-Hans')).toContain('你好');
  expect(await service.translateSubject('你好，你怎么样？', 'en')).toBe(english);
  expect(global.fetch.mock.calls.filter(([url]) => url.pathname === '/languages')).toHaveLength(1);
  expect(JSON.parse(global.fetch.mock.calls[1][1].body)).toEqual({ text: english, source: 'en', target: 'zh' });
  expect(global.fetch.mock.calls.every(([url]) => url.hostname === 'argos.test')).toBe(true);
});

test('mail and Word preserve markup and escape local plain-text output', async () => {
  mockLocal();
  const service = createTranslator(config);
  const html = `<p><span id="r0">${english}</span><span id="r1">已有中文内容。</span></p>`;
  const [word] = await service.translateWord([html], 'zh-Hans');
  expect(word).toBe('<p><span id="r0">你好，你怎么样？ &lt;安全&gt;</span><span id="r1">已有中文内容。</span></p>');
  const mail = await service.translate(`<b>${english}</b><img src="cid:photo"><span translate="no">Keep</span>`, 'zh-Hans');
  expect(mail).toContain('<b>你好，你怎么样？ &lt;安全&gt;</b><img src="cid:photo">');
  expect(mail).toContain('>Keep</span>');
  for (const [url, options] of global.fetch.mock.calls.filter(([url]) => url.pathname === '/translate')) {
    expect(url.hostname).toBe('argos.test');
    expect(JSON.parse(options.body).text).not.toContain('<span');
  }
});

test.each(['http', 'network', 'timeout', 'empty', 'malformed', 'routes'])('falls back to Azure on local %s failure', async kind => {
  global.fetch = jest.fn(async (url) => {
    if (url.hostname === 'azure.test') return json([{ translations: [{ text: 'Azure 译文' }] }]);
    if (url.pathname === '/languages') return json(kind === 'routes' ? {} : routes);
    if (kind === 'network') throw new Error('offline');
    if (kind === 'timeout') throw Object.assign(new Error('timeout'), { name: 'TimeoutError' });
    if (kind === 'http') return { ok: false, status: 503 };
    return json(kind === 'empty' ? { translated_text: ' ', source: 'en', target: 'zh' } : {});
  });
  expect(await createTranslator(config).translateSubject(english, 'zh-Hans')).toBe('Azure 译文');
  expect(global.fetch.mock.calls.at(-1)[0].hostname).toBe('azure.test');
});

test.each(['fr', 'zh-Hant'])('unsupported target %s goes to Azure without local translation', async target => {
  global.fetch = jest.fn(async url => url.pathname === '/languages' ? json(routes) : json([{ translations: [{ text: 'translated' }] }]));
  expect(await createTranslator(config).translateSubject(english, target)).toBe('translated');
  expect(global.fetch.mock.calls).toHaveLength(2);
  expect(global.fetch.mock.calls[1][0].hostname).toBe('azure.test');
});

test('local-only errors never contact Azure', async () => {
  global.fetch = jest.fn(async () => { throw new Error('offline'); });
  await expect(createTranslator({ ...config, localTranslatorFallback: false }).translateSubject(english, 'zh-Hans')).rejects.toThrow('offline');
  expect(global.fetch).toHaveBeenCalledTimes(1);
});

test('detects language offline and preserves already translated Chinese', async () => {
  mockLocal();
  const service = createTranslator({ ...config, localTranslatorFallback: false });
  expect(await service.detect(`<p>${english}</p>`)).toMatchObject({ language: 'en' });
  expect(await service.detect('<p>你好，你怎么样？</p>')).toMatchObject({ language: 'zh-Hans' });
  expect(global.fetch).not.toHaveBeenCalled();
  expect(await service.translateSubject('你好，你怎么样？', 'zh-Hans')).toBe('你好，你怎么样？');
  expect(global.fetch).toHaveBeenCalledTimes(1);
});

test('local calls are counted while offline detection and route discovery do not add translation usage', async () => {
  mockLocal();
  const records = [];
  await context.run({ analytics: { record: row => records.push(row) } }, async () => {
    const service = createTranslator({ ...config, localTranslatorFallback: false });
    await service.detect(english);
    await service.translateSubject(english, 'zh-Hans');
  });
  expect(records).toHaveLength(1);
  expect(records[0]).toMatchObject({ layer: 'upstream', endpoint: 'translate', success: true, units: english.length });
});

test('uncertain offline detection uses Azure detection when fallback is enabled', async () => {
  global.fetch = jest.fn(async () => json([{ language: 'en', score: 1 }]));
  expect(await createTranslator(config).detect('<p>Hello</p>')).toEqual({ language: 'en', score: 1 });
  expect(global.fetch.mock.calls[0][0].hostname).toBe('azure.test');
  expect(global.fetch.mock.calls[0][0].pathname).toBe('/detect');
});

test('long subjects are split without broken surrogate pairs or lost text', async () => {
  const source = ('Hello, how are you? 😀 ').repeat(500);
  global.fetch = jest.fn(async (url, options) => {
    if (url.pathname === '/languages') return json(routes);
    const body = JSON.parse(options.body);
    expect(body.text.length).toBeLessThanOrEqual(4500);
    expect(body.text).not.toMatch(/[\uD800-\uDBFF]$/);
    return json({ translated_text: body.text, source: body.source, target: body.target });
  });
  expect(await createTranslator(config).translateSubject(source, 'zh-Hans')).toBe(source);
  expect(global.fetch.mock.calls.filter(([url]) => url.pathname === '/translate').length).toBeGreaterThan(1);
});

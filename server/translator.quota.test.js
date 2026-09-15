const { createTranslator } = require('./translator');
const originalFetch = global.fetch;
afterEach(() => { global.fetch = originalFetch; jest.useRealTimers(); });
const service = () => createTranslator({ translatorKey: 'test', translatorEndpoint: 'https://translator.example/' });
test('F0 exhaustion blocks subsequent upstream requests and retries in the next UTC month', async () => {
  jest.useFakeTimers(); jest.setSystemTime(new Date('2026-09-15T00:00:00Z'));
  global.fetch = jest.fn(async () => ({ ok: false, status: 403, json: async () => ({ error: { code: 403001 } }) }));
  const translator = service();
  await expect(translator.translateSubject('Hello', 'zh-Hans')).rejects.toMatchObject({ code: 'cloud_quota_exhausted', status: 403 });
  await expect(translator.detect('<p>Hello</p>')).rejects.toThrow('本月云翻译额度已用完');
  expect(global.fetch).toHaveBeenCalledTimes(1);
  jest.setSystemTime(new Date('2026-10-01T00:00:00Z'));
  global.fetch.mockResolvedValue({ ok: true, json: async () => [{ translations: [{ text: '你好' }] }] });
  expect(await translator.translateSubject('Hello', 'zh-Hans')).toBe('你好');
  expect(global.fetch).toHaveBeenCalledTimes(2);
});
test.each([[429, 429000, '过于频繁'], [403, 403000, '认证失败'], [401, 401000, '认证失败']])(
  'ordinary HTTP %s / code %s does not disable cloud translation', async (status, code, message) => {
    global.fetch = jest.fn(async () => ({ ok: false, status, json: async () => ({ error: { code } }) }));
    const translator = service();
    await expect(translator.translateSubject('Hello', 'zh-Hans')).rejects.toThrow(message);
    await expect(translator.translateSubject('Hello', 'zh-Hans')).rejects.toThrow(message);
    expect(global.fetch).toHaveBeenCalledTimes(2);
  });

const { createTranslator } = require('./translator');
const { context } = require('./analytics');
const originalFetch = global.fetch;
afterEach(() => { global.fetch = originalFetch; });
async function capture(run) {
  const records = [];
  const service = createTranslator({ translatorKey: 'key', translatorEndpoint: 'https://translator.invalid' });
  await context.run({ tenantId: 't', userOid: 'u', host: 'word', clientType: 'online', analytics: { record: r => records.push(r) } }, () => run(service));
  return records;
}
test('one logical translation counts every real batch with submitted UTF-16 units', async () => {
  global.fetch = jest.fn(async (_url, init) => ({ ok: true, json: async () => JSON.parse(init.body).map(({ Text }) => ({ translations: [{ text: Text }] })) }));
  const records = await capture(service => service.translate(`<p>${'x'.repeat(90001)}</p>`, 'en'));
  expect(records).toHaveLength(3);
  expect(records.reduce((s,r) => s + r.units, 0)).toBe(90001);
});
test('successful HTTP with broken business markers remains upstream success through repair', async () => {
  global.fetch = jest.fn(async url => ({ ok: true, json: async () => [{ translations: [{ text: url.searchParams.has('textType') ? '<p>broken</p>' : 'translated' }] }] }));
  const records = await capture(service => service.translateWord(['<p><span id="r0">Hello</span></p>'], 'en'));
  expect(records).toHaveLength(2); expect(records.every(r => r.success)).toBe(true);
});
test.each(['http', 'network', 'protocol'])('counts %s failures once', async kind => {
  global.fetch = jest.fn(async () => { if (kind === 'network') throw new Error('network'); return { ok: kind !== 'http', status: 429, json: async () => [] }; });
  const records = await capture(service => expect(service.detect('<p>A😀</p>')).rejects.toThrow());
  expect(records).toHaveLength(1); expect(records[0]).toMatchObject({ endpoint: 'detect', success: false, units: 3 });
});

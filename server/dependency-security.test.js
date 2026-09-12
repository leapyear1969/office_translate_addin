const { createRequire } = require('module');
const { CryptoProvider } = require('@azure/msal-node');

test('MSAL Node generates correlation IDs through its patched CommonJS uuid dependency', () => {
  const uuid = createRequire(require.resolve('@azure/msal-node'))('uuid');
  const id = new CryptoProvider().createNewGuid();
  expect(uuid.validate(id)).toBe(true);
  expect(uuid.version(id)).toBe(4);
  // Exercise the advisory boundary on the exact package resolved by MSAL.
  expect(() => uuid.v5('test', uuid.v5.DNS, new Uint8Array(8), 4)).toThrow(RangeError);
});

test.each(['express', 'body-parser'])('%s resolves qs with both advisory fixes', parent => {
  const qs = createRequire(require.resolve(parent))('qs');
  expect(() => qs.parse('a[0]=1,2,3', { comma: true, arrayLimit: 2, throwOnLimitExceeded: true }))
    .toThrow(RangeError);
  expect(() => qs.stringify({ x: { constructor: { isBuffer: 'invalid' } } })).not.toThrow();
});

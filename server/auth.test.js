const { SignJWT, generateKeyPair } = require('jose');
const { verifyIdentity } = require('./auth');
const config = { tenantId: 'tenant-1', clientId: 'client-1', resource: 'api://localhost/client-1' };
let keys;
beforeAll(async () => { keys = await generateKeyPair('RS256'); });
async function token(overrides = {}) {
  return new SignJWT({ tid: config.tenantId, oid: 'user-1', scp: 'access_as_user', ...overrides })
    .setProtectedHeader({ alg: 'RS256' }).setIssuer('https://issuer.example/tenant-1')
    .setAudience(config.clientId).setIssuedAt().setExpirationTime('5m').sign(keys.privateKey);
}
test('accepts signed identity scoped to the configured application', async () => {
  const result = await verifyIdentity(await token(), config, keys.publicKey, 'https://issuer.example/tenant-1');
  expect(result.oid).toBe('user-1');
});
test.each([{ tid: 'other' }, { scp: 'User.Read' }, { oid: '' }])('rejects unauthorized claims %j', async claims => {
  await expect(verifyIdentity(await token(claims), config, keys.publicKey, 'https://issuer.example/tenant-1')).rejects.toThrow();
});
test('rejects a valid signature intended for another app', async () => {
  await expect(verifyIdentity(await token(), { ...config, clientId: 'other', resource: 'other' }, keys.publicKey, 'https://issuer.example/tenant-1')).rejects.toThrow();
});
test('rejects forged and expired tokens', async () => {
  const other = await generateKeyPair('RS256');
  await expect(verifyIdentity(await token(), config, other.publicKey, 'https://issuer.example/tenant-1')).rejects.toThrow();
  const expired = await new SignJWT({ tid: config.tenantId, oid: 'u', scp: 'access_as_user' }).setProtectedHeader({ alg: 'RS256' }).setIssuer('https://issuer.example/tenant-1').setAudience(config.clientId).setExpirationTime(1).sign(keys.privateKey);
  await expect(verifyIdentity(expired, config, keys.publicKey, 'https://issuer.example/tenant-1')).rejects.toThrow();
});

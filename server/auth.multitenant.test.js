jest.mock('jose', () => ({ ...jest.requireActual('jose'), createRemoteJWKSet: jest.fn() }));
jest.mock('@azure/msal-node', () => ({ ConfidentialClientApplication: jest.fn() }));
const { SignJWT, generateKeyPair, createRemoteJWKSet } = require('jose');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { createAuth } = require('./auth');
const home = '11111111-1111-1111-1111-111111111111';
const customer = '22222222-2222-2222-2222-222222222222';
const config = { tenantId: home, clientId: 'app', resource: 'api://app', clientSecret: 'test',
  authority: 'https://login.partner.microsoftonline.cn', graphBase: 'https://microsoftgraph.chinacloudapi.cn' };
let keys, obo;
const originalFetch = global.fetch;
beforeAll(async () => { keys = await generateKeyPair('RS256'); });
beforeEach(() => {
  createRemoteJWKSet.mockReturnValue(keys.publicKey);
  obo = jest.fn().mockResolvedValue({ accessToken: 'graph-token' });
  ConfidentialClientApplication.mockImplementation(() => ({ acquireTokenOnBehalfOf: obo }));
  global.fetch = jest.fn(async url => {
    if (url.includes('/.well-known/')) {
      const tenant = new URL(url).pathname.split('/')[1];
      return { ok: true, json: async () => ({ issuer: `${config.authority}/${tenant}/v2.0`,
        jwks_uri: `${config.authority}/common/discovery/v2.0/keys` }) };
    }
    return { ok: true, json: async () => ({ id: 'user', displayName: 'Customer' }) };
  });
});
afterEach(() => { global.fetch = originalFetch; });
async function signed(tid = customer, overrides = {}, signingKey = keys.privateKey) {
  return new SignJWT({ tid, oid: 'user', scp: 'access_as_user', ...overrides })
    .setProtectedHeader({ alg: 'RS256' }).setIssuer(`${config.authority}/${tid}/v2.0`)
    .setAudience(config.clientId).setIssuedAt().setExpirationTime('5m').sign(signingKey);
}
test('accepts home and customer tenants and uses each verified tenant for OBO', async () => {
  const auth = createAuth(config);
  for (const tid of [home, customer]) {
    const token = await signed(tid);
    const identity = await auth.authenticate(token);
    expect(identity.tid).toBe(tid);
    expect((await auth.profile(token, identity)).tenantId).toBe(tid);
    expect(obo).toHaveBeenLastCalledWith(expect.objectContaining({ authority: `${config.authority}/${tid}` }));
  }
});
test('rejects issuer/tenant mismatch and forged customer tokens', async () => {
  const auth = createAuth(config);
  await expect(auth.authenticate(await signed(home, { tid: customer }))).rejects.toMatchObject({ status: 401 });
  const other = await generateKeyPair('RS256');
  await expect(auth.authenticate(await signed(customer, {}, other.privateKey))).rejects.toMatchObject({ status: 401 });
});
test('rejects malformed tenant before any discovery request', async () => {
  await expect(createAuth(config).authenticate(await signed('../common'))).rejects.toMatchObject({ status: 401 });
  expect(global.fetch).not.toHaveBeenCalled();
});
test('reports customer consent requirement separately', async () => {
  obo.mockRejectedValue({ errorCode: 'invalid_grant', subError: 'consent_required' });
  await expect(createAuth(config).profile('assertion', { tid: customer, oid: 'user' }))
    .rejects.toMatchObject({ status: 403, code: 'consent_required' });
});

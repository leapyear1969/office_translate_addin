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
test('organization lookup uses the configured cloud and caches names per verified tenant', async () => {
  const original = global.fetch;
  global.fetch = jest.fn(async (url, options) => url.includes('/organization?')
    ? { ok: true, json: async () => ({ value: [{ id: customer, displayName: 'Customer Company' }, { id: home, displayName: 'Home Company' }] }) }
    : original(url, options));
  const auth = createAuth(config);
  for (const tid of [customer, customer, home]) {
    expect((await auth.profile('assertion', { tid, oid: 'user' })).organizationName).toBe(tid === customer ? 'Customer Company' : 'Home Company');
  }
  const calls = global.fetch.mock.calls.filter(([url]) => url.includes('/organization?'));
  expect(calls).toHaveLength(2);
  expect(calls[0][0]).toBe(`${config.graphBase}/v1.0/organization?$select=id,displayName`);
  expect(calls[0][1].headers.Authorization).toBe('Bearer graph-token');
});

test.each(['failure', 'mismatch'])('organization %s does not break login or attach another tenant name', async mode => {
  const original = global.fetch;
  global.fetch = jest.fn(async (url, options) => {
    if (!url.includes('/organization?')) return original(url, options);
    if (mode === 'failure') throw new Error('Graph unavailable');
    return { ok: true, json: async () => ({ value: [{ id: home, displayName: 'Wrong Company' }] }) };
  });
  const profile = await createAuth(config).profile('assertion', { tid: customer, oid: 'user' });
  expect(profile.tenantId).toBe(customer);
  expect(profile.organizationName).toBeUndefined();
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

test('returns an account photo when available and tolerates a missing photo', async () => {
  const profileFetch = global.fetch;
  global.fetch = jest.fn(async (url, options) => {
    if (url.endsWith('/photos/48x48/$value')) return {
      ok: true, headers: new Headers({ 'content-type': 'image/jpeg' }),
      arrayBuffer: async () => Buffer.from('photo'),
    };
    return profileFetch(url, options);
  });
  const auth = createAuth(config);
  expect((await auth.profile('assertion', { tid: customer, oid: 'user' })).photo)
    .toBe('data:image/jpeg;base64,cGhvdG8=');
  global.fetch.mockImplementation(async (url, options) => url.endsWith('/photos/48x48/$value')
    ? { ok: false, status: 404 } : profileFetch(url, options));
  expect((await auth.profile('assertion', { tid: customer, oid: 'user' })).photo).toBeUndefined();
});

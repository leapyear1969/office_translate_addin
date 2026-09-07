jest.mock('@azure/msal-node', () => ({ ConfidentialClientApplication: jest.fn() }));
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { createHash } = require('crypto');
const { createConsent } = require('./consent');
const config = { origin: 'https://addin.example', clientSecret: 'secret', clientId: 'app',
  authority: 'https://login.partner.microsoftonline.cn', graphBase: 'https://microsoftgraph.chinacloudapi.cn' };
const identity = { tid: 'customer', oid: 'user' };
let getAuthCodeUrl, acquireTokenByCode;
beforeEach(() => {
  getAuthCodeUrl = jest.fn(async () => 'https://login.partner.microsoftonline.cn/customer/oauth2/v2.0/authorize');
  acquireTokenByCode = jest.fn(async () => ({ idTokenClaims: { ...identity, nonce: getAuthCodeUrl.mock.calls[0][0].nonce } }));
  ConfidentialClientApplication.mockImplementation(() => ({ getAuthCodeUrl, acquireTokenByCode }));
});
afterEach(() => jest.useRealTimers());
test('uses customer authority, PKCE, matching identity and a one-time callback', async () => {
  const consent = createConsent(config);
  const session = await consent.start(identity);
  expect(ConfidentialClientApplication).toHaveBeenCalledWith(expect.objectContaining({
    auth: expect.objectContaining({ authority: `${config.authority}/customer` }) }));
  expect(session.url).toBe(`${config.origin}/consent.html#state=${session.state}`);
  consent.launch(session.state);
  expect(() => consent.launch(session.state)).toThrow();
  await expect(consent.complete({ state: session.state, code: 'code' })).resolves.toContain('result=success');
  const request = acquireTokenByCode.mock.calls[0][0];
  expect(getAuthCodeUrl.mock.calls[0][0]).toMatchObject({ prompt: 'consent', codeChallengeMethod: 'S256',
    codeChallenge: createHash('sha256').update(request.codeVerifier).digest('base64url') });
  expect(request.redirectUri).toBe(`${config.origin}/auth/consent/callback`);
  await expect(consent.complete({ state: session.state, code: 'code' })).rejects.toThrow();
});
test.each([{ oid: 'other' }, { tid: 'other' }, { nonce: 'wrong' }])('rejects mismatched callback identity %j', async overrides => {
  const consent = createConsent(config);
  const session = await consent.start(identity);
  consent.launch(session.state);
  acquireTokenByCode.mockResolvedValue({ idTokenClaims: { ...identity, nonce: getAuthCodeUrl.mock.calls[0][0].nonce, ...overrides } });
  await expect(consent.complete({ state: session.state, code: 'code' })).resolves.toContain('result=account_mismatch');
});
test('cancellation does not redeem a code and unknown or expired state is rejected', async () => {
  jest.useFakeTimers();
  const consent = createConsent(config);
  const session = await consent.start(identity);
  consent.launch(session.state);
  await expect(consent.complete({ state: session.state, error: 'access_denied' })).resolves.toContain('result=cancelled');
  expect(acquireTokenByCode).not.toHaveBeenCalled();
  await expect(consent.complete({ state: 'unknown', code: 'code' })).rejects.toThrow();
  const expired = await consent.start(identity);
  jest.advanceTimersByTime(10 * 60000 + 1);
  expect(() => consent.launch(expired.state)).toThrow();
});

const originalFetch = global.fetch;
const jwt = (expiresInSeconds = 3600) => `header.${Buffer.from(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + expiresInSeconds })).toString('base64url')}.signature`;
beforeEach(() => {
  jest.resetModules();
  (global as any).OfficeRuntime = { auth: { getAccessToken: jest.fn().mockResolvedValue('sso-token') } };
});
afterEach(() => { global.fetch = originalFetch; delete (global as any).OfficeRuntime; });

test.each([13004, '13004'])('reports SSO resource configuration error %s before calling the backend', async code => {
  (global as any).OfficeRuntime.auth.getAccessToken.mockRejectedValue({ code });
  global.fetch = jest.fn();
  const { authenticate } = await import('./api');
  await expect(authenticate()).rejects.toThrow('页面地址与 WebApplicationInfo/Resource 的域名和端口是否一致');
  expect(global.fetch).not.toHaveBeenCalled();
});
test('preserves structured consent failure and the SSO token for the authorized start endpoint', async () => {
  global.fetch = jest.fn().mockResolvedValue({ ok: false, json: async () => ({ code: 'consent_required', error: '请授权' }) });
  const { authenticate } = await import('./api');
  await expect(authenticate()).rejects.toMatchObject({ code: 'consent_required', token: 'sso-token' });
});
test('ordinary API failures do not carry an SSO token into the consent path', async () => {
  global.fetch = jest.fn().mockResolvedValue({ ok: false, json: async () => ({ error: '服务不可用' }) });
  const { authenticate } = await import('./api');
  const error = await authenticate().catch(error => error);
  expect(error.message).toBe('服务不可用');
  expect(error.token).toBeUndefined();
});

test('concurrent callers share authentication and profile without upgrading silent prompts', async () => {
  let resolve!: (token: string) => void;
  const getToken = (global as any).OfficeRuntime.auth.getAccessToken;
  getToken.mockReturnValueOnce(new Promise(done => { resolve = done; }));
  global.fetch = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ id: 'user' }) });
  const { authenticate } = await import('./api');
  const first = authenticate(false);
  const second = authenticate(true);
  await Promise.resolve();
  expect(getToken).toHaveBeenCalledTimes(1);
  expect(getToken).toHaveBeenCalledWith({ allowSignInPrompt: false, allowConsentPrompt: false });
  resolve('token');
  expect(await first).toEqual(await second);
  expect(global.fetch).toHaveBeenCalledTimes(1);
  await authenticate();
  expect(getToken).toHaveBeenCalledTimes(2);
});

test.each([13013, '13013'])('throttling %s blocks new requests and backs off repeated failures', async code => {
  jest.useFakeTimers();
  try {
    const getToken = (global as any).OfficeRuntime.auth.getAccessToken;
    getToken.mockRejectedValue({ code });
    const { authenticate } = await import('./api');
    await expect(authenticate()).rejects.toMatchObject({ code: '13013' });
    await expect(authenticate()).rejects.toThrow('60 秒');
    expect(getToken).toHaveBeenCalledTimes(1);
    jest.advanceTimersByTime(60000);
    await expect(authenticate()).rejects.toThrow('120 秒');
    expect(getToken).toHaveBeenCalledTimes(2);
    jest.advanceTimersByTime(60000);
    await expect(authenticate()).rejects.toMatchObject({ code: '13013' });
    expect(getToken).toHaveBeenCalledTimes(2);
  } finally { jest.useRealTimers(); }
});

test('timeout does not unlock a native SSO request that is still running', async () => {
  jest.useFakeTimers();
  try {
    let resolve!: (token: string) => void;
    const getToken = (global as any).OfficeRuntime.auth.getAccessToken;
    getToken.mockReturnValue(new Promise(done => { resolve = done; }));
    global.fetch = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ id: 'user' }) });
    const { authenticate } = await import('./api');
    const first = expect(authenticate()).rejects.toThrow('超时');
    await jest.advanceTimersByTimeAsync(20000);
    await first;
    const second = authenticate();
    await Promise.resolve();
    expect(getToken).toHaveBeenCalledTimes(1);
    resolve('token');
    await expect(second).resolves.toMatchObject({ token: 'token' });
  } finally { jest.useRealTimers(); }
});

test('sequential translation authentication reuses successful sign-in and profile', async () => {
  const token = jwt();
  const getToken = (global as any).OfficeRuntime.auth.getAccessToken.mockResolvedValue(token);
  global.fetch = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ id: 'user' }) });
  const { authenticate } = await import('./api');
  await authenticate(false);
  for (let index = 0; index < 10; index++) await authenticate(false);
  expect(getToken).toHaveBeenCalledTimes(1);
  expect(global.fetch).toHaveBeenCalledTimes(1);
});

test.each([120, 3600])('refreshes at expiry safety margin or five-minute limit (lifetime %s)', async lifetime => {
  jest.useFakeTimers();
  try {
    const getToken = (global as any).OfficeRuntime.auth.getAccessToken.mockImplementation(async () => jwt(lifetime));
    global.fetch = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ id: 'user' }) });
    const { authenticate } = await import('./api');
    await authenticate();
    jest.advanceTimersByTime(Math.min(lifetime - 60, 300) * 1000 - 1000);
    await authenticate(false);
    expect(getToken).toHaveBeenCalledTimes(1);
    jest.advanceTimersByTime(1000);
    await authenticate(false);
    expect(getToken).toHaveBeenCalledTimes(2);
  } finally { jest.useRealTimers(); }
});

test('sign-out clears reuse and prevents pending sign-in from restoring the session', async () => {
  const getToken = (global as any).OfficeRuntime.auth.getAccessToken.mockResolvedValue(jwt());
  global.fetch = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ id: 'user' }) });
  const { authenticate, clearAuthentication } = await import('./api');
  await authenticate();
  clearAuthentication();
  let resolve!: (token: string) => void;
  getToken.mockReturnValueOnce(new Promise(done => { resolve = done; }));
  const pending = authenticate();
  await Promise.resolve();
  clearAuthentication();
  resolve(jwt());
  await expect(pending).rejects.toThrow('登录状态已变更');
  await authenticate();
  expect(getToken).toHaveBeenCalledTimes(3);
});

test('401 invalidates the matching session without automatically replaying API work', async () => {
  const getToken = (global as any).OfficeRuntime.auth.getAccessToken.mockResolvedValue(jwt());
  global.fetch = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ id: 'user' }) });
  const { authenticate, api } = await import('./api');
  const session = await authenticate();
  (global.fetch as jest.Mock).mockResolvedValueOnce({ ok: false, status: 401, json: async () => ({ error: 'expired' }) });
  await expect(api('/api/translate', session.token, {})).rejects.toThrow('expired');
  expect(global.fetch).toHaveBeenCalledTimes(2);
  await authenticate();
  expect(getToken).toHaveBeenCalledTimes(2);
});

test.each(['opaque', 'header.e30.signature'])('does not cache tokens without a usable expiry: %s', async token => {
  const getToken = (global as any).OfficeRuntime.auth.getAccessToken.mockResolvedValue(token);
  global.fetch = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ id: 'user' }) });
  const { authenticate } = await import('./api');
  await authenticate();
  await authenticate();
  expect(getToken).toHaveBeenCalledTimes(2);
});

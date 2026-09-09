const originalFetch = global.fetch;
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

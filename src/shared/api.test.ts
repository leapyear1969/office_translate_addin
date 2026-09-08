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

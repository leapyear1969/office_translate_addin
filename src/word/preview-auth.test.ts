/** @jest-environment jsdom */
import { authenticate } from '../shared/api';
import { translatePreview } from './preview';

test('selection changes keep translating without repeating native SSO or profile requests', async () => {
  const originalFetch = global.fetch;
  const token = `header.${btoa(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + 3600 }))}.signature`;
  const getAccessToken = jest.fn().mockResolvedValue(token);
  (global as any).OfficeRuntime = { auth: { getAccessToken } };
  global.fetch = jest.fn().mockImplementation(async path => ({ ok: true,
    json: async () => path === '/api/me' ? { id: 'user' } : { html: '<div>译文</div>' },
  }));
  try {
    await authenticate(false);
    for (let index = 0; index < 10; index++) {
      await expect(translatePreview(`Selection ${index}`, 'zh-Hans')).resolves.toBe('译文');
    }
    expect(getAccessToken).toHaveBeenCalledTimes(1);
    expect((global.fetch as jest.Mock).mock.calls.filter(([path]) => path === '/api/me')).toHaveLength(1);
    expect((global.fetch as jest.Mock).mock.calls.filter(([path]) => path === '/api/translate')).toHaveLength(10);
  } finally {
    global.fetch = originalFetch;
    delete (global as any).OfficeRuntime;
  }
});

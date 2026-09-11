/** @jest-environment jsdom */
import { readFileSync } from 'fs';
import { join } from 'path';
let mockSignedIn = true;
jest.mock('@azure/msal-browser', () => ({
  PublicClientApplication: class {
    async initialize() {} async handleRedirectPromise() { return null; }
    getActiveAccount() { return mockSignedIn ? { username: 'admin@example.invalid' } : null; }
    getAllAccounts() { return []; } setActiveAccount() {}
    async acquireTokenSilent() { return { accessToken: 'test-token' }; }
    async loginRedirect() {} async logoutRedirect() {}
  }, InteractionRequiredAuthError: class extends Error {},
}));
const originalFetch = global.fetch;
const now = Date.now();
const today = new Date(now + 8 * 3600000).toISOString().slice(0, 10);
const payload = {
  coverage: { started_at: String(now - 30 * 86400000), retained_from: '2020-01-01' }, health: {},
  summary: { users: 1, requests: 5 }, daily: [{ date: today, requests: 5 }], activeDaily: { [today]: 1 },
  platforms: [{ host: 'word', client_type: 'online', users: 1, requests: 5 }],
  totalUsers: 1, pageSize: 50, users: [{ display_name: '<img src=x onerror=alert(1)>', mail: 'person@example.invalid', tenant_id: 't', user_oid: 'u', first_used_at: now, last_used_at: now, days: 1, requests: 5, platforms: 'word / online' }],
};
const settle = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };
beforeEach(() => {
  jest.resetModules(); mockSignedIn = true; history.replaceState(null, '', '/admin/usage');
  document.documentElement.innerHTML = readFileSync(join(__dirname, 'admin.html'), 'utf8');
  (AbortSignal as any).timeout ||= () => new AbortController().signal;
  global.fetch = jest.fn(async (url: any) => ({ ok: true, json: async () => url === '/admin/config'
    ? { configured: true, scope: 'api://test/access_as_user', clientId: 'test', redirectUri: 'http://localhost/admin/usage' }
    : url === '/api/admin/session' ? { tenants: ['t'] } : payload })) as any;
});
afterEach(() => { global.fetch = originalFetch; });
test('renders real report response safely, sends filters, and hides stale data on query failure', async () => {
  await import('./admin'); await settle();
  expect(document.getElementById('report')!.hidden).toBe(false);
  expect(document.getElementById('cards')!.textContent).toContain('期间使用人数1');
  expect(document.querySelector('#users img')).toBeNull();
  expect(document.getElementById('users')!.textContent).toContain('<img src=x onerror=alert(1)>');
  (document.getElementById('client-type') as HTMLSelectElement).value = 'local';
  document.getElementById('filters')!.dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  expect((global.fetch as jest.Mock).mock.calls.at(-1)[0]).toContain('client_type=local');
  (global.fetch as jest.Mock).mockResolvedValueOnce({ ok: false, status: 503, json: async () => ({ error: '统计暂不可用' }) });
  document.getElementById('filters')!.dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  expect(document.getElementById('report')!.hidden).toBe(true);
  expect(document.getElementById('status')!.textContent).toBe('统计暂不可用');
});
test('configuration missing and signed-out states do not fetch administrator data', async () => {
  (global.fetch as jest.Mock).mockResolvedValueOnce({ ok: true, json: async () => ({ configured: false }) });
  await import('./admin'); await settle();
  expect(document.getElementById('status')!.textContent).toContain('尚未配置');
  expect(global.fetch).toHaveBeenCalledTimes(1);
  expect((document.getElementById('login') as HTMLButtonElement).disabled).toBe(true);
});
test('history before collection is unavailable rather than zero', async () => {
  (global.fetch as jest.Mock).mockImplementation(async (url: string) => ({ ok: true, json: async () => url === '/admin/config'
    ? { configured: true, scope: 's' } : url === '/api/admin/session' ? { tenants: ['t'] }
      : { ...payload, coverage: { ...payload.coverage, started_at: String(now + 86400000) } } }));
  await import('./admin'); await settle();
  expect(document.getElementById('cards')!.textContent).toBe('');
  expect(document.getElementById('trend')!.textContent).toContain('历史数据不可用');
});
test('API page separates layers and does not show provider units on business rows', async () => {
  history.replaceState(null, '', '/admin/api');
  (global.fetch as jest.Mock).mockImplementation(async (url: string) => ({ ok: true, json: async () => url === '/admin/config'
    ? { configured: true, scope: 's' } : url === '/api/admin/session' ? { tenants: ['t'] }
      : { ...payload, totals: [{ layer: 'business', requests: 1 }, { layer: 'upstream', requests: 3 }],
        interfaces: [{ layer: 'business', endpoint: 'translate', requests: 1, successes: 1, failures: 0 }],
        daily: [], platforms: [], users: [] } }));
  await import('./admin'); await settle();
  expect(document.getElementById('cards')!.textContent).toContain('后端业务请求1');
  expect(document.getElementById('cards')!.textContent).toContain('上游实际请求3');
  expect(document.getElementById('interfaces')!.textContent).toContain('不重复计量');
});

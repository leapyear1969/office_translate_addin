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
const apiPayload = {
  ...payload, totals: [{ layer: 'business', requests: 1 }, { layer: 'upstream', requests: 8 }],
  interfaces: [{ layer: 'business', endpoint: 'translate', requests: 1, successes: 1, failures: 0 }],
  daily: [], platforms: [], users: [], totalUsers: 2, pageSize: 1,
  topUsers: [
    { tenant_id: 't', user_oid: 'u', display_name: '<img src=x onerror=alert(1)>', mail: 'person@example.invalid', submitted_units: 12345, requests: 7, hosts: ['outlook', 'word'], client_types: ['local', 'online'] },
    { tenant_id: 't', user_oid: 'unknown-user', display_name: null, mail: null, submitted_units: 20, requests: 1, hosts: ['unknown'], client_types: ['unknown'] },
  ],
  topTenants: [{ tenant_id: 't', submitted_units: 12365, requests: 8, hosts: ['outlook', 'unknown', 'word'], client_types: ['local', 'online', 'unknown'] }],
};
function apiPage(report: unknown = apiPayload) {
  history.replaceState(null, '', '/admin/api');
  (global.fetch as jest.Mock).mockImplementation(async (url: string) => ({ ok: true, json: async () => url === '/admin/config'
    ? { configured: true, scope: 's' } : url === '/api/admin/session' ? { tenants: ['t'] } : report }));
}
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

test('settings URL denies readers and exposes the navigation only to managers', async () => {
  history.replaceState(null, '', '/admin/settings');
  await import('./admin'); await settle();
  expect(document.getElementById('settings-link')!.hidden).toBe(true);
  expect(document.getElementById('admin-settings')!.hidden).toBe(true);
  expect(document.getElementById('status')!.textContent).toContain('没有管理员配置权限');
  expect((global.fetch as jest.Mock).mock.calls.some(([url]) => url === '/api/admin/settings')).toBe(false);
});

test('global administrator defaults to all tenants and can filter a tenant', async () => {
  (global.fetch as jest.Mock).mockImplementation(async (url: string) => ({ ok: true, json: async () => url === '/admin/config'
    ? { configured: true, scope: 's' } : url === '/api/admin/session' ? { tenants: ['*'] } : payload }));
  await import('./admin'); await settle();
  const tenant = document.querySelector<HTMLInputElement>('input#tenant')!;
  expect(tenant.placeholder).toContain('所有租户');
  expect(new URL((global.fetch as jest.Mock).mock.calls.at(-1)[0], 'http://localhost').searchParams.get('tenant')).toBe('');
  tenant.value = '11111111-1111-1111-1111-111111111111';
  document.getElementById('filters')!.dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  expect((global.fetch as jest.Mock).mock.calls.at(-1)[0]).toContain(`tenant=${tenant.value}`);
});
test('renders real report response safely, sends filters, and hides stale data on query failure', async () => {
  await import('./admin'); await settle();
  expect(document.getElementById('report')!.hidden).toBe(false);
  expect(document.getElementById('cards')!.textContent).toContain('期间使用人数1');
  expect(document.getElementById('top-users-panel')!.hidden).toBe(true);
  expect(document.getElementById('top-tenants-panel')!.hidden).toBe(true);
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
test('API rankings display text totals, upstream requests and all platforms safely', async () => {
  apiPage(); await import('./admin'); await settle();
  expect(document.getElementById('top-users-panel')!.hidden).toBe(false);
  expect(document.getElementById('top-tenants-panel')!.hidden).toBe(false);
  const cells = (selector: string) => Array.from(document.querySelectorAll(selector)).map(cell => cell.textContent);
  expect(cells('#top-users thead th')).toEqual(['排名', '姓名', '邮箱', '租户 / 账号', '提交文本量（UTF-16）', '上游请求数', '平台', '客户端']);
  expect(cells('#top-users tbody tr:first-child td')).toEqual(['1', '<img src=x onerror=alert(1)>', 'person@example.invalid', 't / u', '12,345', '7', 'Outlook、Word', '桌面客户端、Web页面']);
  expect(cells('#top-users tbody tr:nth-child(2) td')).toEqual(['2', '未获取', '未获取', 't / unknown-user', '20', '1', '未识别', '未识别']);
  expect(cells('#top-tenants thead th')).toEqual(['排名', '租户 ID', '提交文本量（UTF-16）', '上游请求数', '平台', '客户端']);
  expect(cells('#top-tenants tbody td')).toEqual(['1', 't', '12,365', '8', 'Outlook、未识别、Word', '桌面客户端、Web页面、未识别']);
  expect(document.querySelector('#top-users img')).toBeNull();
  expect(document.querySelectorAll('#top-users tbody tr:first-child .numeric.centered')).toHaveLength(3);
  expect(document.querySelectorAll('#top-tenants tbody .numeric.centered')).toHaveLength(3);
  expect(document.getElementById('top-tenants-panel')!.textContent).toContain('用户搜索');
});
test('API rankings share report filters and remain present when the user list is paged', async () => {
  apiPage(); await import('./admin'); await settle();
  const original = document.getElementById('top-users')!.textContent;
  document.getElementById('next')!.click(); await settle();
  expect(new URL((global.fetch as jest.Mock).mock.calls.at(-1)[0], 'http://localhost').searchParams.get('page')).toBe('2');
  expect(document.getElementById('top-users')!.textContent).toBe(original);
  for (const [id, value] of [['from', today], ['to', today], ['tenant', 't'], ['host', 'word'], ['client-type', 'local'], ['search', 'person']]) {
    (document.getElementById(id) as HTMLInputElement).value = value;
  }
  document.getElementById('filters')!.dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  const url = new URL((global.fetch as jest.Mock).mock.calls.at(-1)[0], 'http://localhost');
  expect(url.pathname).toBe('/api/admin/api-usage');
  expect(Object.fromEntries(url.searchParams)).toEqual({ from: today, to: today, tenant: 't', host: 'word', client_type: 'local', search: 'person', page: '1' });
  expect(document.getElementById('top-users')!.textContent).toBe(original);
});
test.each([undefined, []])('API rankings handle missing or empty arrays (%j)', async rankings => {
  apiPage({ ...apiPayload, topUsers: rankings, topTenants: rankings });
  await import('./admin'); await settle();
  for (const id of ['top-users', 'top-tenants']) {
    expect(document.getElementById(id)!.textContent).toBe('所选范围内暂无已观测记录');
    expect(document.querySelector(`#${id} table`)).toBeNull();
  }
  expect(document.getElementById('report')!.hidden).toBe(false);
});
test('API rankings clear previous data for unavailable history and hide on query failure', async () => {
  apiPage(); await import('./admin'); await settle();
  expect(document.querySelector('#top-users tbody')).not.toBeNull();
  (global.fetch as jest.Mock).mockResolvedValueOnce({ ok: true, json: async () => ({ ...apiPayload, coverage: { ...apiPayload.coverage, started_at: String(now + 86400000) } }) });
  document.getElementById('filters')!.dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  for (const id of ['top-users', 'top-tenants']) {
    expect(document.getElementById(id)!.textContent).toBe('历史数据不可用');
    expect(document.querySelector(`#${id} table`)).toBeNull();
  }
  (global.fetch as jest.Mock).mockResolvedValueOnce({ ok: false, status: 503, json: async () => ({ error: '统计暂不可用' }) });
  document.getElementById('filters')!.dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  expect(document.getElementById('report')!.hidden).toBe(true);
  expect(document.getElementById('status')!.textContent).toBe('统计暂不可用');
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

import type { PublicClientApplication } from '@azure/msal-browser';

type Row = Record<string, any>;
const el = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const input = (id: string) => el<HTMLInputElement>(id);
const isApi = location.pathname === '/admin/api';
const isSettings = location.pathname === '/admin/settings';
let msal: PublicClientApplication;
let InteractionRequiredAuthError: typeof import('@azure/msal-browser').InteractionRequiredAuthError;
let CacheLookupPolicy: typeof import('@azure/msal-browser').CacheLookupPolicy;
let scope = '';
let page = 1;
let generation = 0;
let currentParams = new URLSearchParams();
const fmt = (value: unknown) => Number(value || 0).toLocaleString('zh-CN');
const time = (value: unknown) => value ? new Date(Number(value)).toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai', hour12: false }) : '未获取';
const date = (value: number) => new Date(value + 8 * 3600000).toISOString().slice(0, 10);
const label = (value: unknown) => ({ word: 'Word', outlook: 'Outlook', online: 'Web页面', local: '桌面客户端', unknown: '未识别', business: '后端业务', upstream: '上游服务' }[String(value)] || String(value || '未获取'));
function status(message = '', error = false) { el('status').textContent = message; el('status').className = error ? 'error' : ''; }
function table(target: string, headers: string[], rows: (string | number)[][], centeredColumns: number[] = []) {
  const container = el(target); container.replaceChildren();
  if (!rows.length) { const p = document.createElement('p'); p.className = 'empty'; p.textContent = '所选范围内暂无已观测记录'; container.append(p); return; }
  const table = document.createElement('table'), head = document.createElement('thead'), tr = document.createElement('tr'), body = document.createElement('tbody');
  for (const [index, title] of headers.entries()) { const th = document.createElement('th'); th.scope = 'col'; th.textContent = title; if (centeredColumns.includes(index)) th.className = 'centered'; tr.append(th); }
  head.append(tr); table.append(head, body);
  for (const row of rows) { const tr = document.createElement('tr'); for (const [index, value] of row.entries()) { const td = document.createElement('td'); td.textContent = typeof value === 'number' ? fmt(value) : value; if (typeof value === 'number') td.className = 'numeric'; if (centeredColumns.includes(index)) td.classList.add('centered'); tr.append(td); } body.append(tr); }
  container.append(table);
}
function cards(items: [string, number, string][]) {
  el('cards').replaceChildren();
  for (const [title, value, note] of items) {
    const card = document.createElement('div'); card.className = 'card';
    for (const [tag, css, text] of [['div', 'label', title], ['div', 'value', fmt(value)], ['small', '', note]]) { const node = document.createElement(tag); node.className = css; node.textContent = text; card.append(node); }
    el('cards').append(card);
  }
}
async function token() {
  const account = msal.getActiveAccount();
  if (!account) throw new Error('请先使用管理员账号登录。');
  // The standalone admin page forbids framing, including MSAL's hidden-iframe fallback.
  // Keep cached tokens and refresh-token renewal; expired sessions use the login button.
  try { return (await msal.acquireTokenSilent({ scopes: [scope], account, cacheLookupPolicy: CacheLookupPolicy.AccessTokenAndRefreshToken })).accessToken; }
  catch (error) { if (error instanceof InteractionRequiredAuthError) throw new Error('登录已过期或需要授权，请点击“重新登录”。'); throw new Error('无法获取访问权限，请重新登录。'); }
}
async function get(path: string, body?: unknown) {
  const response = await fetch(path, { method: body === undefined ? 'GET' : 'PUT', body: body === undefined ? undefined : JSON.stringify(body), headers: { Authorization: `Bearer ${await token()}`, ...(body === undefined ? {} : { 'Content-Type': 'application/json' }) }, cache: 'no-store', signal: AbortSignal.timeout(15000) });
  if (!response.ok) {
    let message = `查询失败（${response.status}），请稍后重试。`;
    try { const data = await response.json(); if (typeof data.error === 'string') message = data.error; } catch { /* Rate limit may return text. */ }
    throw new Error(message);
  }
  return response.json();
}
function trend(data: Row, from: string, to: string) {
  const container = el('trend'); container.replaceChildren();
  const values = new Map<string, number[]>();
  for (const row of data.daily) {
    const v = values.get(row.date) || [0, 0];
    if (isApi) v[row.layer === 'business' ? 0 : 1] = row.requests;
    else { v[0] = row.requests; v[1] = data.activeDaily[row.date] || 0; }
    values.set(row.date, v);
  }
  const start = Math.max(Date.parse(from), Date.parse(date(Number(data.coverage.started_at))), Date.parse(data.coverage.retained_from));
  for (let d = start; d <= Date.parse(to); d += 86400000) { const key = new Date(d).toISOString().slice(0, 10); if (!values.has(key)) values.set(key, [0, 0]); }
  const max = Math.max(1, ...[...values.values()].map(v => Math.max(...v)));
  if (!values.size) { container.textContent = '该范围尚未开始采集，无可用趋势。'; return; }
  for (const [day, counts] of [...values.entries()].sort(([a], [b]) => a.localeCompare(b))) {
    const row = document.createElement('div'); row.className = 'trend-row';
    const dateLabel = document.createElement('span'); dateLabel.textContent = day;
    const bars = document.createElement('div'); bars.className = 'bars';
    counts.forEach((n, i) => { const bar = document.createElement('div'); bar.className = `bar${i ? ' secondary' : ''}`; bar.style.width = `${n / max * 100}%`; if (!n) bar.style.visibility = 'hidden'; bars.append(bar); });
    const text = document.createElement('span'); text.textContent = isApi ? `后端 ${fmt(counts[0])} / 上游 ${fmt(counts[1])}` : `${fmt(counts[0])} 请求 / ${fmt(counts[1])} 人`;
    row.append(dateLabel, bars, text); container.append(row);
  }
}
function render(data: Row) {
  const coverage = data.coverage;
  const gap = data.health.lastGapAt || coverage.last_gap_at;
  el('coverage').textContent = `采集起点：${time(coverage.started_at)}。日汇总保留至 ${coverage.retained_from} 起；采集前及已清理的历史不可用。${gap ? ` 已知统计缺口：${time(gap)} 附近存在写入异常，报表可能不完整，空白日期不代表确认未使用。` : '进程中断可能造成未观测请求，统计不是审计账本。'}`;
  if (currentParams.get('to')! < date(Number(coverage.started_at))) {
    el('report').hidden = false;
    el('cards').replaceChildren(); el('trend').textContent = '所选日期早于采集起点，历史数据不可用。';
    for (const id of ['interfaces', 'platforms', 'users', 'top-users', 'top-tenants']) { el(id).textContent = '历史数据不可用'; }
    el('user-count').textContent = ''; el('page').textContent = '';
    el<HTMLButtonElement>('previous').disabled = el<HTMLButtonElement>('next').disabled = true;
    return;
  }
  if (isApi) {
    const find = (layer: string) => data.totals.find((r: Row) => r.layer === layer)?.requests || 0;
    cards([['后端业务请求', find('business'), '检测与翻译，含自动预检'], ['上游实际请求', find('upstream'), '每次真实尝试，含分批与修复']]);
    const labels = (values: string[] = []) => values.map(label).join('、') || '未获取';
    table('top-users', ['排名', '姓名', '邮箱', '租户 / 账号', '提交文本量（UTF-16）', '上游请求数', '平台', '客户端'], (data.topUsers || []).map((r: Row, i: number) => [i + 1, r.display_name || '未获取', r.mail || '未获取', `${r.tenant_id} / ${r.user_oid}`, r.submitted_units, r.requests, labels(r.hosts), labels(r.client_types)]), [0, 4, 5]);
    table('top-tenants', ['排名', '租户 ID', '组织名称', '提交文本量（UTF-16）', '上游请求数', '平台', '客户端'], (data.topTenants || []).map((r: Row, i: number) => [i + 1, r.tenant_id, r.organization_name || '未获取', r.submitted_units, r.requests, labels(r.hosts), labels(r.client_types)]), [0, 3, 4]);
    table('interfaces', ['层级', '接口', '请求数', '成功', '失败', '提交文本量（UTF-16）'], data.interfaces.map((r: Row) => [label(r.layer), r.endpoint, r.requests, r.successes, r.failures, r.layer === 'upstream' ? r.submitted_units : '不重复计量']), [2, 3, 4, 5]);
    table('platforms', ['平台', '客户端', '层级', '请求数'], data.platforms.map((r: Row) => [label(r.host), label(r.client_type), label(r.layer), r.requests]), [3]);
    table('users', ['姓名', '邮箱', '租户 / 账号', '后端请求', '上游请求'], data.users.map((r: Row) => [r.display_name || '未获取', r.mail || '未获取', `${r.tenant_id} / ${r.user_oid}`, r.business, r.upstream]), [3, 4]);
  } else {
    cards([['期间使用人数', data.summary.users, '按已验证账号去重'], ['翻译请求次数', data.summary.requests, '包含自动预览刷新与失败请求']]);
    table('platforms', ['平台', '客户端', '使用人数', '翻译请求'], data.platforms.map((r: Row) => [label(r.host), label(r.client_type), r.users, r.requests]), [2, 3]);
    table('users', ['姓名', '邮箱', '租户 / 账号', '首次观测使用', '最近观测使用', '使用天数', '翻译请求', '期间平台 / 客户端'], data.users.map((r: Row) => [r.display_name || '未获取', r.mail || '未获取', `${r.tenant_id} / ${r.user_oid}`, time(r.first_used_at), time(r.last_used_at), r.days, r.requests, String(r.platforms || '').split(', ').map(platform => platform.split(' / ').map(label).join(' / ')).join(', ')]), [5, 6]);
  }
  trend(data, currentParams.get('from')!, currentParams.get('to')!);
  el('user-count').textContent = `${fmt(data.totalUsers)} 位用户`;
  el('page').textContent = `第 ${page} / ${Math.max(1, Math.ceil(data.totalUsers / data.pageSize))} 页`;
  el<HTMLButtonElement>('previous').disabled = page <= 1;
  el<HTMLButtonElement>('next').disabled = page * data.pageSize >= data.totalUsers;
  el('report').hidden = false;
}
async function load() {
  const gen = ++generation;
  status('正在查询…'); el('report').hidden = true;
  currentParams.set('page', String(page));
  try { const data = await get(`/api/admin/${isApi ? 'api-usage' : 'overview'}?${currentParams}`); if (gen !== generation) return; render(data); status(); }
  catch (error) { if (gen === generation) status(error instanceof Error ? error.message : '查询暂不可用。', true); }
}
function collect() {
  currentParams = new URLSearchParams();
  for (const [id, key] of [['from', 'from'], ['to', 'to'], ['tenant', 'tenant'], ['host', 'host'], ['client-type', 'client_type'], ['search', 'search']]) currentParams.set(key, input(id).value);
}
async function main() {
  el(isSettings ? 'settings-link' : isApi ? 'api-link' : 'usage-link').setAttribute('aria-current', 'page');
  el('page-title').textContent = isSettings ? '管理员设置' : isApi ? 'API 用量' : '用户使用';
  el('page-description').textContent = isSettings ? '管理后台访问权限，无需修改服务器配置。' : isApi ? '区分后端请求与上游调用，查看实际请求规模。' : '查看谁在使用，以及来自哪个平台与客户端。';
  if (isSettings) { document.querySelector<HTMLElement>('.badge')!.hidden = true; document.querySelector<HTMLElement>('footer')!.hidden = true; }
  for (const id of ['interfaces-panel', 'top-users-panel', 'top-tenants-panel']) el(id).hidden = !isApi;
  el('platform-note').textContent = isApi ? '两个层级独立统计，不相加。提交文本量不是计费字符数。' : '同一用户可出现在多个分组中，总人数按账号去重。';
  el('trend-note').textContent = isApi ? '绿色：后端 · 灰蓝：上游' : '绿色：翻译请求 · 灰蓝：使用人数';
  const now = Date.now(); input('from').value = date(now - 29 * 86400000); input('to').value = date(now);
  input('from').max = input('to').max = date(now);
  const config = await fetch('/admin/config', { cache: 'no-store' }).then(r => { if (!r.ok) throw new Error('无法读取管理站配置。'); return r.json(); });
  if (!config.configured) { status('管理站尚未配置浏览器登录。请由维护人员完成 ADMIN_CLIENT_ID 与 API 权限配置。'); el<HTMLButtonElement>('login').disabled = true; return; }
  scope = config.scope;
  const { PublicClientApplication, InteractionRequiredAuthError: AuthError, CacheLookupPolicy: LookupPolicy } = await import(/* webpackChunkName: "admin-auth" */ '@azure/msal-browser');
  CacheLookupPolicy = LookupPolicy;
  InteractionRequiredAuthError = AuthError;
  msal = new PublicClientApplication({ auth: { clientId: config.clientId, authority: config.authority, redirectUri: config.redirectUri, navigateToLoginRequestUrl: false }, cache: { cacheLocation: 'sessionStorage' }, system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } } });
  await msal.initialize();
  const result = await msal.handleRedirectPromise();
  if (result?.account) msal.setActiveAccount(result.account);
  else if (!msal.getActiveAccount() && msal.getAllAccounts().length === 1) msal.setActiveAccount(msal.getAllAccounts()[0]);
  el('login').onclick = () => { void msal.loginRedirect({ scopes: [scope], prompt: 'select_account' }).catch(() => status('无法启动登录，请稍后重试。', true)); };
  el('logout').onclick = () => { el('workspace').hidden = true; el('admin-settings').hidden = true; void msal.logoutRedirect({ postLogoutRedirectUri: config.redirectUri }).catch(() => status('退出失败，请重试。', true)); };
  if (!msal.getActiveAccount()) { status('请使用获授权的管理员账号登录。'); return; }
  el('account').textContent = msal.getActiveAccount()?.username || '';
  el('login').textContent = '重新登录'; el('logout').hidden = false;
  const access = await get('/api/admin/session');
  el('settings-link').hidden = !access.canManageAdmins;
  if (isSettings) {
    if (!access.canManageAdmins) { status('此账号没有管理员配置权限，请联系后台权限管理员。', true); return; }
    const { initSettings } = await import('./settings');
    await initSettings(get, status); return;
  }
  if (access.retained_from) {
    input('from').min = input('to').min = access.retained_from;
    if (input('from').value < access.retained_from) input('from').value = access.retained_from;
  }
  if (access.tenants.includes('*')) {
    const tenant = document.createElement('input'); tenant.id = 'tenant'; tenant.placeholder = '所有租户（可输入租户 ID 筛选）';
    tenant.setAttribute('aria-label', '租户 ID，留空查询所有租户'); el('tenant').replaceWith(tenant);
  } else {
    for (const tenant of access.tenants) { const option = document.createElement('option'); option.value = tenant; option.textContent = tenant; el('tenant').append(option); }
  }
  el('workspace').hidden = false;
  el('filters').onsubmit = event => { event.preventDefault(); page = 1; collect(); void load(); };
  el('previous').onclick = () => { page--; void load(); }; el('next').onclick = () => { page++; void load(); };
  collect(); await load();
}
void main().catch(error => status(error instanceof Error ? error.message : '管理站加载失败，请刷新重试。', true));

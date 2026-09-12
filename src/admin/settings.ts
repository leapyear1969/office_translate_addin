type Admin = { tenant_id: string; email?: string; user_oid?: string; tenants: string[]; can_manage_admins?: boolean };
type Settings = { revision: number; entries: Admin[]; selfKeys: string[] };
const keyOf = (e: Admin) => `${e.tenant_id}:${e.user_oid ? `oid:${e.user_oid}` : `email:${e.email}`}`;
const guidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}';

export async function initSettings(request: (path: string, body?: unknown) => Promise<Settings>, status: (message?: string, error?: boolean) => void) {
  const el = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
  const container = el('admin-entries');
  let saved: Settings;
  let dirty = false;
  const markDirty = () => { dirty = true; el('settings-state').textContent = '有未保存的更改'; };
  const fields = el<HTMLFieldSetElement>('settings-fields');
  const addRow = (entry: Admin, self = false) => {
    const row = document.createElement('fieldset'); row.className = 'admin-entry panel';
    const legend = document.createElement('legend'); legend.textContent = self ? '管理员 · 当前账号' : '管理员'; row.append(legend);
    const grid = document.createElement('div'); grid.className = 'settings-grid'; row.append(grid);
    function field(title: string, name: string, value: string, options?: [string, string][]) {
      const label = document.createElement('label'); label.textContent = title;
      const control = options ? document.createElement('select') : document.createElement('input');
      control.name = name;
      for (const [value, text] of options || []) { const option = document.createElement('option'); option.value = value; option.textContent = text; control.append(option); }
      control.value = value; label.append(control); grid.append(label); return control;
    }
    const tenant = field('账号登录所属租户 ID', 'tenant_id', entry.tenant_id) as HTMLInputElement;
    tenant.required = true; tenant.pattern = guidPattern; tenant.placeholder = '例如 00000000-0000-0000-0000-000000000000';
    const identity = field('用户对象 ID', 'identity', entry.user_oid || '') as HTMLInputElement;
    identity.required = true; identity.maxLength = 36; identity.pattern = guidPattern;
    identity.placeholder = '用户对象 ID（GUID）';
    const scope = field('可查看的租户', 'scope', entry.tenants.includes('*') ? 'all' : 'selected', [['selected', '指定租户'], ['all', '所有租户（含以后新增）']]);
    const tenants = field('指定租户 ID（逗号或空格分隔）', 'tenants', entry.tenants.filter(t => t !== '*').join(', ')) as HTMLInputElement;
    tenants.maxLength = 10000;
    const syncScope = () => { tenants.parentElement!.hidden = scope.value === 'all'; tenants.required = scope.value !== 'all'; };
    scope.onchange = syncScope; syncScope();
    const role = field('后台角色', 'role', entry.can_manage_admins ? 'manager' : 'reader', [['reader', '查询管理员 · 仅查看报表'], ['manager', '权限管理员 · 可管理所有管理员']]);
    if (self) { tenant.readOnly = true; identity.readOnly = true; role.disabled = true; }
    const note = document.createElement('p'); note.className = 'muted'; note.textContent = self ? '当前账号的登录身份和权限管理资格不能在此移除。' : '权限管理员可以调整所有账号的权限，包括授予所有租户的访问权。'; row.append(note);
    if (!entry.user_oid && entry.email) note.textContent = `旧邮箱条目 ${entry.email} 已停用。请核对原获批账号，填写该租户内的对象 ID，或移除此条目。`;
    const remove = document.createElement('button'); remove.type = 'button'; remove.className = 'remove-admin'; remove.textContent = '移除管理员'; remove.disabled = self;
    remove.onclick = () => { row.remove(); markDirty(); el('add-admin').focus(); };
    row.append(remove); container.append(row);
    return tenant;
  };
  const render = (data: Settings) => {
    saved = data; container.replaceChildren();
    for (const entry of data.entries) addRow(entry, data.selfKeys.includes(keyOf(entry)));
    dirty = false; el('settings-state').textContent = `${data.entries.length} 位管理员 · 已保存`;
  };
  const reload = async () => {
    fields.disabled = true; status('正在读取管理员配置…');
    try { render(await request('/api/admin/settings')); el('admin-settings').hidden = false; status(); }
    catch (error) { status(error instanceof Error ? error.message : '读取失败，请重试。', true); }
    finally { fields.disabled = false; }
  };
  el('admin-settings-form').oninput = markDirty;
  el('admin-settings-form').onchange = markDirty;
  el('add-admin').onclick = () => { addRow({tenant_id: '', user_oid: '', tenants: [], can_manage_admins: false}).focus(); markDirty(); };
  el('reload-admins').onclick = () => { void reload(); };
  el('admin-settings-form').onsubmit = async event => {
    event.preventDefault();
    if (!saved) return;
    const entries: Admin[] = Array.from(container.querySelectorAll<HTMLFieldSetElement>('.admin-entry')).map(row => {
      const value = (name: string) => row.querySelector<HTMLInputElement | HTMLSelectElement>(`[name="${name}"]`)!.value.trim();
      return {tenant_id:value('tenant_id'), user_oid:value('identity'), tenants:value('scope') === 'all' ? ['*'] : value('tenants').split(/[,，\s]+/).filter(Boolean), can_manage_admins:value('role') === 'manager'};
    });
    fields.disabled = true; status('正在保存管理员配置…');
    try { render(await request('/api/admin/settings', {revision:saved.revision,entries})); status('管理员配置已保存，权限立即生效。'); }
    catch (error) { status(error instanceof Error ? error.message : '保存失败，请重新加载确认配置。', true); }
    finally { fields.disabled = false; }
  };
  window.addEventListener('beforeunload', event => { if (dirty) { event.preventDefault(); event.returnValue = ''; } });
  await reload();
}

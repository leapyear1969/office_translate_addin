const { PGlite } = require('@electric-sql/pglite');
const { openStore } = require('./analytics-store');
const { createAnalytics } = require('./analytics');
const { createApp } = require('./app');
const tid = '11111111-1111-1111-1111-111111111111', oid = '22222222-2222-2222-2222-222222222222';
const readerOid = '33333333-3333-3333-3333-333333333333';
const owner = {tenant_id:tid,user_oid:oid,tenants:['*']};
let engine,store,analytics,server;
beforeEach(async()=>{
  engine = new PGlite();
  store = await openStore('test',Date.now(),{connect:async()=>{},query:(sql,args)=>engine.query(sql,args),end:async()=>{}});
  analytics = createAnalytics('test',async()=>store);
  const app = createApp({origin:'https://test.invalid',analyticsAdmins:[owner]}, {
    analytics,authenticate:async token=>({tid,oid:token==='owner'?oid:readerOid,email:'owner@example.invalid'}),
    profile:async(_token,id)=>({id:id.oid,tenantId:id.tid,mail:id.oid===oid?'owner@example.invalid':'reader@example.invalid'}),
  });
  server = await new Promise(resolve=>{const s=app.listen(0,'127.0.0.1',()=>resolve(s));});
});
afterEach(async()=>{server.closeAllConnections();await new Promise(resolve=>server.close(resolve));await analytics.close();await engine.close();});
const request=(path='/settings',token='owner',body,origin)=>fetch(`http://127.0.0.1:${server.address().port}/api/admin${path}`,{
  method:body?'PUT':'GET',headers:{...(token?{Authorization:`Bearer ${token}`}:{ }),'Content-Type':'application/json',...(origin?{Origin:origin}:{})},body:body?JSON.stringify(body):undefined,
});

test('permission changes persist minimal audit diffs with verified actor and reject rewrites',async()=>{
  const initial=await (await request()).json();
  const reader={tenant_id:tid,user_oid:readerOid,tenants:[tid],can_manage_admins:false};
  const write=(revision,entries)=>request('/settings','owner',{revision,entries,actor:{tenant_id:'spoof',user_oid:'spoof'},token:'never-store-this'});
  expect((await write(1,[...initial.entries,reader])).status).toBe(200);
  expect((await write(2,[...initial.entries,{...reader,can_manage_admins:true}])).status).toBe(200);
  expect((await write(3,initial.entries)).status).toBe(200);
  expect((await write(3,initial.entries)).status).toBe(409);
  const rows=(await engine.query('SELECT * FROM usage_analytics.admin_settings_audit ORDER BY revision')).rows;
  expect(rows.map(r=>[r.previous_revision,r.revision])).toEqual([[1,2],[2,3],[3,4]]);
  for(const row of rows) {
    expect(row).toMatchObject({actor_tenant_id:tid,actor_user_oid:oid});
    expect(Number.isFinite(Date.parse(row.changed_at))).toBe(true);
  }
  expect(rows[0].changes).toEqual([{before:null,after:reader}]);
  expect(rows[1].changes).toEqual([{before:reader,after:{...reader,can_manage_admins:true}}]);
  expect(rows[2].changes).toEqual([{before:{...reader,can_manage_admins:true},after:null}]);
  expect(JSON.stringify(rows)).not.toMatch(/spoof|never-store-this|Bearer/);
  for(const sql of ['UPDATE usage_analytics.admin_settings_audit SET actor_user_oid=\'spoof\'',
    'DELETE FROM usage_analytics.admin_settings_audit','TRUNCATE usage_analytics.admin_settings_audit']) {
    await expect(engine.query(sql)).rejects.toThrow('append-only');
  }
});

test('audit failure rolls back permission changes and revision',async()=>{
  const initial=await (await request()).json();
  await engine.query(`ALTER TABLE usage_analytics.admin_settings_audit ADD CONSTRAINT simulate_failure CHECK (revision < 0)`);
  await expect(store.saveAdminSettings({revision:1,entries:[{...initial.entries[0],tenants:[tid]}],
    actor:{tenant_id:tid,user_oid:oid}})).rejects.toThrow();
  expect(await store.adminSettings([])).toEqual({revision:1,entries:initial.entries});
  expect((await engine.query('SELECT * FROM usage_analytics.admin_settings_audit')).rows).toHaveLength(0);
  await expect(store.saveAdminSettings({revision:1,entries:initial.entries})).rejects.toThrow('actor');
});

test('legacy email manager stays inactive while a bound owner can migrate it',async()=>{
  const legacy={tenant_id:tid,email:'reader@example.invalid',tenants:['*'],can_manage_admins:true};
  await store.adminSettings([{...owner,can_manage_admins:true},legacy]);
  expect((await request('/settings','reader')).status).toBe(403);
  const initial=await (await request()).json();
  const bound={tenant_id:tid,user_oid:readerOid,tenants:['*'],can_manage_admins:true};
  expect((await request('/settings','owner',{revision:1,entries:[initial.entries[0],bound]})).status).toBe(200);
  expect((await request('/settings','reader')).status).toBe(200);
});

test('empty bootstrap stays closed and scoped bootstrap remains read-only by default',async()=>{
  expect(await store.adminSettings([])).toEqual({revision:0,entries:[]});
  expect((await store.adminSettings([{...owner,tenants:[tid],can_manage_admins:false}])).revision).toBe(1);
  expect((await request('/settings')).status).toBe(403);
  expect(await (await request('/session')).json()).toMatchObject({tenants:[tid],canManageAdmins:false});
});

test('bootstrap once, save, immediate authorization and revocation, persistence and no settings leakage',async()=>{
  expect((await request('/settings','')).status).toBe(401);
  expect((await request('/settings','reader')).status).toBe(403);
  const initial=await (await request()).json();
  expect(initial.entries[0].can_manage_admins).toBe(true);
  const reader={tenant_id:tid,user_oid:readerOid,tenants:[tid],can_manage_admins:false};
  const saved=await (await request('/settings','owner',{revision:initial.revision,entries:[...initial.entries,reader]})).json();
  expect(saved.revision).toBe(2);
  expect(await (await request('/session','reader')).json()).toMatchObject({tenants:[tid],canManageAdmins:false});
  expect((await request('/settings','reader')).status).toBe(403);
  expect((await request('/settings','reader',{revision:2,entries:[{...reader,can_manage_admins:true}]})).status).toBe(403);
  expect((await request('/overview?tenant='+oid,'reader')).status).toBe(403);
  const report=await (await request('/overview','reader')).json();
  expect(JSON.stringify(report)).not.toContain('can_manage_admins');
  expect((await store.adminSettings([{...owner,tenants:[tid]}])).entries).toEqual(saved.entries);
  expect((await request('/settings','owner',{revision:2,entries:initial.entries})).status).toBe(200);
  expect((await request('/session','reader')).status).toBe(403);
  await analytics.close();
  analytics=createAnalytics('test',async()=>store);
  expect((await analytics.adminSettings([owner])).revision).toBe(3);
  expect((await analytics.adminSettings([owner])).entries).toEqual(initial.entries);
});

test('rejects invalid input, self lockout, duplicates, foreign origins and stale writes',async()=>{
  const initial=await (await request()).json();
  const write=entries=>request('/settings','owner',{revision:1,entries});
  expect((await write([])).status).toBe(400);
  expect((await write([{...owner,can_manage_admins:false}])).status).toBe(400);
  expect((await write([{...owner,user_oid:readerOid,can_manage_admins:true}])).status).toBe(400);
  expect((await write([initial.entries[0],initial.entries[0]])).status).toBe(400);
  expect((await write([{...owner,tenant_id:'bad',can_manage_admins:true}])).status).toBe(400);
  expect((await write([{...owner,can_manage_admins:'true'}])).status).toBe(400);
  expect((await request('/settings','owner',{revision:1,entries:initial.entries},'https://evil.invalid')).status).toBe(403);
  const responses=await Promise.all([write(initial.entries),write(initial.entries)]);
  expect(responses.map(r=>r.status).sort()).toEqual([200,409]);
  expect((await request()).status).toBe(200);
  await expect(analytics.saveAdminSettings({revision:1,entries:initial.entries})).rejects.toMatchObject({status:409});
  expect((await analytics.adminSettings([owner])).revision).toBe(2);
});

test('new email grants are rejected and storage errors fail closed',async()=>{
  const initial=await (await request()).json();
  const manager={tenant_id:tid,email:'reader@example.invalid',tenants:['*'],can_manage_admins:true};
  expect((await request('/settings','owner',{revision:1,entries:[...initial.entries,manager]})).status).toBe(400);
  expect((await request('/settings','reader')).status).toBe(403);
  const spy=jest.spyOn(store,'adminSettings').mockRejectedValueOnce(new Error('private database details'));
  const warn=jest.spyOn(console,'warn').mockImplementation(()=>{});
  const failed=await request();
  expect(failed.status).toBe(503);expect(await failed.text()).not.toContain('private database');
  spy.mockRestore();warn.mockRestore();
});

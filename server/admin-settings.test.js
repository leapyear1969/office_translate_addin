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

test('empty bootstrap stays closed and scoped bootstrap remains read-only by default',async()=>{
  expect(await store.adminSettings([])).toEqual({revision:0,entries:[]});
  expect((await store.adminSettings([{...owner,tenants:[tid],can_manage_admins:false}])).revision).toBe(1);
  expect((await request('/settings')).status).toBe(403);
  expect(await (await request('/session')).json()).toEqual({tenants:[tid],canManageAdmins:false});
});

test('bootstrap once, save, immediate authorization and revocation, persistence and no settings leakage',async()=>{
  expect((await request('/settings','')).status).toBe(401);
  expect((await request('/settings','reader')).status).toBe(403);
  const initial=await (await request()).json();
  expect(initial.entries[0].can_manage_admins).toBe(true);
  const reader={tenant_id:tid,user_oid:readerOid,tenants:[tid],can_manage_admins:false};
  const saved=await (await request('/settings','owner',{revision:initial.revision,entries:[...initial.entries,reader]})).json();
  expect(saved.revision).toBe(2);
  expect(await (await request('/session','reader')).json()).toEqual({tenants:[tid],canManageAdmins:false});
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

test('email grants use verified profile, support management and fail closed on storage errors',async()=>{
  const initial=await (await request()).json();
  const manager={tenant_id:tid,email:'reader@example.invalid',tenants:['*'],can_manage_admins:true};
  expect((await request('/settings','owner',{revision:1,entries:[...initial.entries,manager]})).status).toBe(200);
  expect((await request('/settings','reader')).status).toBe(200);
  const spy=jest.spyOn(store,'adminSettings').mockRejectedValueOnce(new Error('private database details'));
  const warn=jest.spyOn(console,'warn').mockImplementation(()=>{});
  const failed=await request();
  expect(failed.status).toBe(503);expect(await failed.text()).not.toContain('private database');
  spy.mockRestore();warn.mockRestore();
});

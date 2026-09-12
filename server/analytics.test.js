const { PGlite } = require('@electric-sql/pglite');
const { openStore,day,cutoff } = require('./analytics-store');
const { createAnalytics,context,recordRequest } = require('./analytics');
let engine,store;
const now=Date.now(),today=day(now);
const base={tenantId:'tenant-a',userOid:'user-a',host:'word',clientType:'online',layer:'business',endpoint:'translate',startedAt:now,success:true,units:0};
const filter={from:day(now-30*86400000),to:today,tenants:['tenant-a'],page:1,pageSize:50};
beforeEach(async()=>{
  engine=new PGlite();
  store=await openStore('test',now-30*86400000,{connect:async()=>{},query:async(sql,args)=>{
    const result=await engine.query(sql,args); return {...result,rowCount:result.affectedRows ?? result.rows.length};
  },end:()=>engine.close()});
});
afterEach(async()=>{await store?.close();});
test('custom retention controls writes, coverage and cleanup while retaining the boundary day',async()=>{
  const custom=await openStore('test',now,{connect:async()=>{},query:(sql,args)=>engine.query(sql,args),end:async()=>{}},2);
  const boundary=cutoff(now,2);
  const boundaryTime=Date.parse(`${boundary}T12:00:00+08:00`);
  await custom.record({...base,startedAt:boundaryTime});
  await custom.record({...base,startedAt:boundaryTime-86400000});
  let result=await custom.query({...filter,from:'2000-01-01'},'usage');
  expect(result.coverage.retained_from).toBe(boundary);
  expect(result.summary.requests).toBe(1);
  // The original store can write older data, simulating a shortened retention setting.
  await store.record({...base,startedAt:boundaryTime-86400000});
  await custom.cleanup(now);
  result=await custom.query({...filter,from:'2000-01-01'},'usage');
  expect(result.summary.requests).toBe(1);
  expect(result.daily[0].date).toBe(boundary);
});
test('deduplicates users and days across platforms; detect-only accounts do not become active',async()=>{
  await store.record(base);await store.record({...base,host:'outlook',clientType:'local',success:false});
  await store.record({...base,userOid:'detect-only',endpoint:'detect'});await store.record({...base,tenantId:'tenant-b'});
  const result=await store.query(filter,'usage');
  expect(result.summary).toEqual({requests:2,users:1});expect(result.users[0]).toMatchObject({days:1,requests:2});
  expect(result.platforms).toHaveLength(2);expect(result.platforms.every(p=>p.users===1)).toBe(true);expect(result.activeDaily[today]).toBe(1);
  expect((await store.query(filter,'api')).totals[0]).toMatchObject({requests:3,successes:2,failures:1});
});
test('keeps layer totals separate, counts UTF-16, and preserves start-day min/max first use',async()=>{
  const earlier=Date.parse(`${today}T00:00:00+08:00`)-60000;
  await store.record(base);await store.record({...base,startedAt:earlier});
  for(let i=0;i<3;i++)await store.record({...base,layer:'upstream',units:'A😀'.length,success:i!==2});
  const result=await store.query(filter,'api');
  expect(result.totals.find(r=>r.layer==='business').requests).toBe(2);
  expect(result.totals.find(r=>r.layer==='upstream')).toMatchObject({requests:3,submitted_units:9,failures:1});
  expect(result.topUsers).toEqual([expect.objectContaining({tenant_id:'tenant-a',user_oid:'user-a',requests:3,submitted_units:9})]);
  expect(result.topTenants).toEqual([expect.objectContaining({tenant_id:'tenant-a',requests:3,submitted_units:9})]);
  expect((await store.query(filter,'usage')).users[0]).toMatchObject({first_used_at:earlier,last_used_at:now,days:2});
});

test('global reports include new tenants while explicit tenant filters remain scoped',async()=>{
  await store.record(base);
  const all = {...filter,tenants:['*'],host:'word'};
  expect((await store.query(all,'usage')).summary.users).toBe(1);
  await store.record({...base,tenantId:'tenant-b'});
  const report=await store.query(all,'usage');
  expect(report.summary).toEqual({requests:2,users:2});
  expect(report.users.map(u=>u.tenant_id).sort()).toEqual(['tenant-a','tenant-b']);
  expect(report.activeDaily[today]).toBe(2);
  expect((await store.query(all,'api')).totals[0].requests).toBe(2);
  expect((await store.query(filter,'usage')).summary.users).toBe(1);
  expect((await store.query({...all,tenants:[]},'usage')).summary.users).toBe(0);
});
test('literal search, pagination, empty grants and tenant isolation apply to report sections',async()=>{
  for(const tenantId of ['tenant-a','tenant-b']){
    await store.record({...base,tenantId});await store.profile({tenantId,userOid:'user-a',displayName:'A% User',mail:'same@example.com'});
  }
  await store.record({...base,userOid:'user-b'});
  const result=await store.query({...filter,search:'%',pageSize:1},'usage');
  expect(result.summary.users).toBe(1);expect(result.users[0].tenant_id).toBe('tenant-a');
  expect((await store.query({...filter,tenants:[]},'usage')).summary.users).toBe(0);
  expect((await store.query({...filter,page:2,pageSize:1},'usage')).users).toHaveLength(1);
});
test('API rankings aggregate upstream text by entity across days, endpoints and platforms',async()=>{
  const upstream={...base,layer:'upstream'};
  await store.record({...upstream,units:60});
  await store.record({...upstream,units:40,host:'outlook',clientType:'local',endpoint:'detect',success:false,startedAt:now-86400000});
  await store.record({...upstream,userOid:'user-b',units:30});
  await store.record({...upstream,tenantId:'tenant-b',units:70,clientType:'local'});
  await store.record({...upstream,tenantId:'tenant-b',userOid:'user-b',units:70,host:'unknown',clientType:'unknown'});
  await store.record({...base,tenantId:'business-only'});
  await store.profile({tenantId:'tenant-a',userOid:'user-a',displayName:'A User',mail:'a@example.invalid'});
  await engine.query("DELETE FROM usage_analytics.users WHERE tenant_id='tenant-b' AND user_oid='user-b'");
  const result=await store.query({...filter,tenants:['*']},'api');
  expect(result.topUsers.map(r=>[r.tenant_id,r.user_oid,r.submitted_units,r.requests])).toEqual([
    ['tenant-a','user-a',100,2],['tenant-b','user-a',70,1],['tenant-b','user-b',70,1],['tenant-a','user-b',30,1],
  ]);
  expect(result.topUsers[0]).toMatchObject({display_name:'A User',mail:'a@example.invalid',hosts:['outlook','word'],client_types:['local','online']});
  expect(result.topUsers[2]).toMatchObject({display_name:null,mail:null,hosts:['unknown'],client_types:['unknown']});
  expect(result.topTenants).toEqual([
    {tenant_id:'tenant-b',organization_name:null,submitted_units:140,requests:2,hosts:['unknown','word'],client_types:['local','unknown']},
    {tenant_id:'tenant-a',organization_name:null,submitted_units:130,requests:3,hosts:['outlook','word'],client_types:['local','online']},
  ]);
  await store.record({...upstream,tenantId:'zero-text',units:0});
  expect((await store.query({...filter,tenants:['zero-text']},'api')).topTenants).toEqual([
    expect.objectContaining({tenant_id:'zero-text',submitted_units:0,requests:1}),
  ]);
});
test('tenant names backfill historical rankings and survive missing profile metadata',async()=>{
  await store.record({...base,layer:'upstream',units:10});
  await store.record({...base,tenantId:'tenant-b',layer:'upstream',units:20});
  await store.profile({...base,organizationName:'Company A'});
  await store.profile({...base,userOid:'another-user'});
  let result=await store.query({...filter,tenants:['*']},'api');
  expect(result.topTenants.map(r=>[r.tenant_id,r.organization_name,r.submitted_units])).toEqual([['tenant-b',null,20],['tenant-a','Company A',10]]);
  await store.profile({...base,organizationName:'Renamed Company'});
  result=await store.query(filter,'api');
  expect(result.topTenants[0].organization_name).toBe('Renamed Company');
});

test('both API rankings use the same tenant grants and filters as the report',async()=>{
  const upstream={...base,layer:'upstream'};
  await store.record({...upstream,units:60});
  await store.record({...upstream,units:40,host:'outlook',clientType:'local',startedAt:now-86400000});
  await store.record({...upstream,userOid:'user-b',units:30});
  await store.record({...upstream,tenantId:'tenant-b',units:70,clientType:'local'});
  await store.profile({tenantId:'tenant-a',userOid:'user-a',displayName:'A% User',mail:'a@example.invalid'});
  const all={...filter,tenants:['*']};
  for(const [extra,users,tenants] of [
    [{tenants:['tenant-a']},[['tenant-a','user-a',100,2],['tenant-a','user-b',30,1]],[['tenant-a',130,3]]],
    [{tenants:[]},[],[]],
    [{from:today,to:today},[['tenant-b','user-a',70,1],['tenant-a','user-a',60,1],['tenant-a','user-b',30,1]],[['tenant-a',90,2],['tenant-b',70,1]]],
    [{host:'word'},[['tenant-b','user-a',70,1],['tenant-a','user-a',60,1],['tenant-a','user-b',30,1]],[['tenant-a',90,2],['tenant-b',70,1]]],
    [{clientType:'local'},[['tenant-b','user-a',70,1],['tenant-a','user-a',40,1]],[['tenant-b',70,1],['tenant-a',40,1]]],
    [{userOid:'user-b'},[['tenant-a','user-b',30,1]],[['tenant-a',30,1]]],
    [{search:'%'},[['tenant-a','user-a',100,2]],[['tenant-a',100,2]]],
    [{search:'missing'},[],[]],
  ]){
    const result=await store.query({...all,...extra},'api');
    expect(result.topUsers.map(r=>[r.tenant_id,r.user_oid,r.submitted_units,r.requests])).toEqual(users);
    expect(result.topTenants.map(r=>[r.tenant_id,r.submitted_units,r.requests])).toEqual(tenants);
  }
});
test('API text rankings are limited to ten, deterministic and independent of user pagination',async()=>{
  for(let i=11;i>=1;i--)await store.record({...base,tenantId:'many',userOid:`u${String(i).padStart(2,'0')}`,layer:'upstream',units:20});
  await store.record({...base,tenantId:'single',layer:'upstream',units:190});
  const all={...filter,tenants:['*'],pageSize:1};
  const first=await store.query(all,'api');
  expect(first.topUsers.map(r=>`${r.tenant_id}/${r.user_oid}`)).toEqual(['single/user-a',...Array.from({length:9},(_,i)=>`many/u${String(i+1).padStart(2,'0')}`)]);
  expect(first.topTenants.map(r=>[r.tenant_id,r.submitted_units,r.requests])).toEqual([['many',220,11],['single',190,1]]);
  const second=await store.query({...all,page:2},'api');
  expect(second.topUsers).toEqual(first.topUsers);expect(second.topTenants).toEqual(first.topTenants);
  expect(second.users).not.toEqual(first.users);
  for(let i=11;i>=1;i--)await store.record({...base,tenantId:`t${String(i).padStart(2,'0')}`,layer:'upstream',units:70});
  await store.record({...base,tenantId:'tie-requests',layer:'upstream',units:35});
  await store.record({...base,tenantId:'tie-requests',layer:'upstream',units:35});
  const tied=await store.query(all,'api');
  expect(tied.topTenants.map(r=>r.tenant_id)).toEqual(['many','single','tie-requests','t01','t02','t03','t04','t05','t06','t07']);
  expect(tied.topUsers.map(r=>r.tenant_id)).toEqual(['single','tie-requests','t01','t02','t03','t04','t05','t06','t07','t08']);
});
test('cleanup preserves first use, deletion removes only the target account, and failed writes roll back',async()=>{
  await store.record(base);await store.record({...base,tenantId:'tenant-b'});
  await expect(store.record({...base,userOid:'rollback',units:null})).rejects.toThrow();
  expect((await engine.query("SELECT * FROM usage_analytics.users WHERE user_oid='rollback'")).rows).toHaveLength(0);
  expect(await store.deleteUser('tenant-a','user-a')).toEqual({usageRows:1,users:1});
  expect((await store.query(filter,'usage')).summary.users).toBe(0);
  expect((await store.query({...filter,tenants:['tenant-b']},'usage')).summary.users).toBe(1);
  await store.cleanup(now+367*86400000);
  expect((await store.query({...filter,tenants:['tenant-b']},'usage')).summary.users).toBe(0);
  expect(Number((await engine.query("SELECT first_used_at FROM usage_analytics.users WHERE tenant_id='tenant-b'")).rows[0].first_used_at)).toBe(now);
});
test('calendar leap day and concurrent request context do not mix identities',async()=>{
  expect(cutoff(Date.parse('2024-02-29T03:00:00Z'))).toBe('2023-02-28');
  const records=[];
  await Promise.all(['a','b'].map(async userOid=>context.run({tenantId:'t',userOid,host:'word',clientType:'local',analytics:{record:r=>records.push(r)}},async()=>{
    await new Promise(resolve=>setImmediate(resolve));recordRequest('upstream','translate',now,true,2);
  })));
  expect(records.map(r=>r.userOid).sort()).toEqual(['a','b']);
  expect(Object.keys(records[0]).sort()).toEqual(['tenantId','userOid','host','clientType','layer','endpoint','startedAt','success','units'].sort());
});
test('async queue serializes transactions and never throws to translation on connection failure',async()=>{
  const target=store;
  const analytics=createAnalytics('test',async()=>({...target,close:async()=>{}}));
  try{for(let i=0;i<20;i++)analytics.record(base);await analytics.flush();expect((await analytics.query(filter,'usage')).summary.requests).toBe(20);}
  finally{await analytics.close();}
  const warning=jest.spyOn(console,'warn').mockImplementation(()=>{});
  const blocked=createAnalytics('test',async()=>{throw new Error('Database unavailable');});
  try{expect(()=>blocked.record(base)).not.toThrow();await expect(blocked.query(filter,'usage')).rejects.toThrow('unavailable');}
  finally{await blocked.close();warning.mockRestore();}
});

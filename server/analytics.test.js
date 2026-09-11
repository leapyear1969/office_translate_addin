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
  expect((await store.query(filter,'usage')).users[0]).toMatchObject({first_used_at:earlier,last_used_at:now,days:2});
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

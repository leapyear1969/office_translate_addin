const { PGlite } = require('@electric-sql/pglite');
const { openStore } = require('./analytics-store');
const { recoverAdmin } = require('../scripts/recover-analytics-admin');
const tenant_id='d7125684-0e28-40b5-aba2-ea9580f2a201';
const target={tenant_id,user_oid:'1fae1516-383a-4f53-9653-324afb9b0616',tenants:['*'],can_manage_admins:true};
const legacy={tenant_id,email:'jason@majun.fun',tenants:['*'],can_manage_admins:true};
let engine,store;
beforeEach(async()=>{
  engine=new PGlite();
  store=await openStore('test',Date.now(),{connect:async()=>{},query:(sql,args)=>engine.query(sql,args),end:async()=>{}});
  await store.adminSettings([legacy]);
});
afterEach(async()=>{await engine.close();});
test('preview does not write; recovery preserves existing entries and records database actor',async()=>{
  expect(await recoverAdmin(engine,[target])).toMatchObject({mode:'preview',revision:1,preservedEntries:1});
  expect((await store.adminSettings([])).entries).toEqual([legacy]);
  expect(await recoverAdmin(engine,[target],true)).toMatchObject({mode:'saved',revision:2});
  expect((await store.adminSettings([])).entries).toEqual([legacy,target]);
  const audit=(await engine.query('SELECT * FROM usage_analytics.admin_recovery_audit')).rows[0];
  expect(audit).toMatchObject({source:'offline_recovery',revision:2,previous_revision:1,changes:{before:[],after:[target]}});
  expect(audit.database_user).toBeTruthy();
  await recoverAdmin(engine,[target],true);
  expect((await store.adminSettings([])).entries).toEqual([legacy,target]);
  await expect(engine.query('DELETE FROM usage_analytics.admin_recovery_audit')).rejects.toThrow('append-only');
});
test('invalid identities and an active serving process fail closed',async()=>{
  for (const entries of [[legacy],[],[target,target],[{...target,can_manage_admins:false}]]) {
    await expect(recoverAdmin(engine,entries,true)).rejects.toThrow();
  }
  const busy={query:(sql,args)=>sql.includes('pg_try_advisory_xact_lock') ? {rows:[{owned:false}]} : engine.query(sql,args)};
  await expect(recoverAdmin(busy,[target],true)).rejects.toThrow('Stop all backend');
  expect(await store.adminSettings([])).toEqual({revision:1,entries:[legacy]});
});
test('audit failure rolls back recovery',async()=>{
  await recoverAdmin(engine,[target],true);
  await engine.query('ALTER TABLE usage_analytics.admin_recovery_audit ADD CONSTRAINT fail_next CHECK (revision < 3)');
  await expect(recoverAdmin(engine,[{...target,tenants:[tenant_id]}],true)).rejects.toThrow();
  expect(await store.adminSettings([])).toEqual({revision:2,entries:[legacy,target]});
});

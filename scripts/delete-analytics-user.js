// Stop all serving processes and drain requests before offline maintenance.
const { readConfig }=require('../server/config');
const { openStore }=require('../server/analytics-store');
const { parseAdmins }=require('../server/admin');
async function main(){
  const args=process.argv.slice(2);
  if(args.length!==3||args[2]!=='--confirm')throw new Error('Usage');
  const [tenantId,userOid]=args;
  parseAdmins(JSON.stringify([{tenant_id:tenantId,user_oid:userOid,tenants:[tenantId]}]));
  const config=readConfig();if(!config.analyticsDatabaseUrl)throw new Error('No connection');
  const store=await openStore(config.analyticsDatabaseUrl,undefined,undefined,config.analyticsRetentionMonths);
  try{const result=await store.deleteUser(tenantId.toLowerCase(),userOid.toLowerCase());console.log(`Deleted ${result.users} user row and ${result.usageRows} daily rows. Future requests will be collected after restart.`);}
  finally{await store.close();}
}
main().catch(()=>{console.error('Deletion not completed. Usage: npm run analytics:delete-user -- <tenant_id> <user_oid> --confirm\nStop all serving processes first. Verify database configuration and account IDs. This permanently removes statistics; there is no undo.');process.exitCode=1;});

// Offline maintenance only: never expose this operation through an HTTP route.
const { parseAdmins } = require('../server/admin');

async function recoverAdmin(client, configured, confirm = false) {
  const entries = parseAdmins(JSON.stringify(configured));
  if (entries.length !== 1 || !entries[0].user_oid || entries[0].can_manage_admins !== true) {
    throw new Error('Set ANALYTICS_ADMINS to exactly one verified object-ID administrator with can_manage_admins:true.');
  }
  const target = entries[0];
  await client.query('BEGIN');
  try {
    // Same lock as the serving process: refuse recovery while it is running.
    const lock = await client.query('SELECT pg_try_advisory_xact_lock(731204, 2) AS owned');
    if (!lock.rows[0].owned) throw new Error('Stop all backend instances and drain requests before recovery.');
    const current = (await client.query('SELECT revision,entries FROM usage_analytics.admin_settings WHERE id=1 FOR UPDATE')).rows[0];
    if (!current) throw new Error('No initialized settings found; restart the backend to import ANALYTICS_ADMINS normally.');
    const matches = e => e.tenant_id.toLowerCase() === target.tenant_id && e.user_oid?.toLowerCase() === target.user_oid;
    const previous = current.entries.filter(matches);
    const next = [...current.entries.filter(e => !matches(e)),target];
    if (next.length > 100) throw new Error('Administrator limit exceeded; review the existing settings first.');
    const summary = {mode:confirm ? 'saved' : 'preview',revision:current.revision,
      target,previous,preservedEntries:next.length-1};
    if (confirm) {
      // Attribute offline changes to the database session, never invent a verified Entra actor.
      await client.query(`CREATE TABLE IF NOT EXISTS usage_analytics.admin_recovery_audit (
        revision INTEGER PRIMARY KEY, previous_revision INTEGER NOT NULL,
        changed_at TIMESTAMPTZ NOT NULL DEFAULT clock_timestamp(),
        database_user TEXT NOT NULL DEFAULT SESSION_USER,
        source TEXT NOT NULL CHECK (source='offline_recovery'), changes JSONB NOT NULL)`);
      await client.query('REVOKE ALL ON usage_analytics.admin_recovery_audit FROM PUBLIC');
      await client.query('DROP TRIGGER IF EXISTS admin_recovery_append_only ON usage_analytics.admin_recovery_audit');
      await client.query(`CREATE TRIGGER admin_recovery_append_only BEFORE UPDATE OR DELETE OR TRUNCATE
        ON usage_analytics.admin_recovery_audit FOR EACH STATEMENT EXECUTE FUNCTION usage_analytics.reject_audit_mutation()`);
      await client.query(`INSERT INTO usage_analytics.admin_recovery_audit
        (revision,previous_revision,source,changes) VALUES ($1,$2,'offline_recovery',$3::jsonb)`,
      [current.revision+1,current.revision,JSON.stringify({before:previous,after:[target]})]);
      const saved = await client.query(`UPDATE usage_analytics.admin_settings SET entries=$1::jsonb,revision=revision+1
        WHERE id=1 AND revision=$2 RETURNING revision`,[JSON.stringify(next),current.revision]);
      if (!saved.rows.length) throw new Error('Settings revision changed; retry after reviewing the preview.');
      summary.revision = saved.rows[0].revision;
    }
    await client.query('COMMIT');
    return summary;
  } catch (error) { await client.query('ROLLBACK'); throw error; }
}

async function main() {
  const args = process.argv.slice(2);
  if (args.length > 1 || (args.length && args[0] !== '--confirm')) throw new Error('Usage: npm run analytics:recover-admin -- [--confirm]');
  const { readConfig } = require('../server/config');
  const { Client } = require('pg');
  const config = readConfig();
  if (!config.analyticsDatabaseUrl) throw new Error('ANALYTICS_DATABASE_URL is required.');
  const client = new Client({connectionString:config.analyticsDatabaseUrl,connectionTimeoutMillis:5000,statement_timeout:5000});
  try {
    await client.connect();
    console.log(JSON.stringify(await recoverAdmin(client,config.analyticsAdmins,args[0] === '--confirm'),null,2));
  } finally { await client.end(); }
}
if (require.main === module) main().catch(error => {
  // Do not print connection strings or database internals from driver errors.
  console.error(error.code ? 'Recovery failed. Check database access/schema and stop all backend instances.' : error.message);
  process.exitCode = 1;
});
module.exports = { recoverAdmin };

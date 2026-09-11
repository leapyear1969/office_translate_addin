const { Client } = require('pg');
const day = time => new Date(time + 8 * 3600000).toISOString().slice(0, 10);
function cutoff(now = Date.now()) {
  const date = new Date(`${day(now)}T00:00:00Z`), d = date.getUTCDate();
  date.setUTCDate(1); date.setUTCMonth(date.getUTCMonth() - 12);
  date.setUTCDate(Math.min(d, new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0)).getUTCDate()));
  return date.toISOString().slice(0, 10);
}

// One connection is serialized by analytics.js. A session lock prevents a
// second serving process or offline maintenance from writing concurrently.
async function openStore(connectionString, now = Date.now(), injectedClient) {
  const client = injectedClient || new Client({ connectionString, connectionTimeoutMillis: 2000,
    statement_timeout: 3000, query_timeout: 4000, idle_in_transaction_session_timeout: 5000, keepAlive: true });
  let broken = false;
  client.on?.('error', () => { broken = true; });
  const q = async (sql, args = []) => { if (broken) throw new Error('Statistics connection lost'); return client.query(sql, args); };
  async function transaction(run, readonly = false) {
    await q(readonly ? 'BEGIN ISOLATION LEVEL REPEATABLE READ READ ONLY' : 'BEGIN');
    try { const result = await run(); await q('COMMIT'); return result; }
    catch (error) { try { await q('ROLLBACK'); } catch { broken = true; } throw error; }
  }
  const meta = async () => Object.fromEntries((await q('SELECT * FROM usage_analytics.meta')).rows.map(r => [r.key, r.value]));
  try {
    await client.connect();
    if (!(await q('SELECT pg_try_advisory_lock(731204, 2) AS owned')).rows[0].owned) throw new Error('Another statistics process owns this database');
    await transaction(async () => {
      await q('CREATE SCHEMA IF NOT EXISTS usage_analytics');
      await q('CREATE TABLE IF NOT EXISTS usage_analytics.meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)');
      const version = (await meta()).schema_version;
      if (version && version !== '1') throw new Error('Unsupported statistics schema');
      await q(`CREATE TABLE IF NOT EXISTS usage_analytics.users (
        tenant_id TEXT NOT NULL, user_oid TEXT NOT NULL, display_name TEXT, mail TEXT,
        first_used_at BIGINT, last_used_at BIGINT, PRIMARY KEY(tenant_id,user_oid))`);
      await q(`CREATE TABLE IF NOT EXISTS usage_analytics.api_usage_daily (
        date TEXT NOT NULL, tenant_id TEXT NOT NULL, user_oid TEXT NOT NULL,
        host TEXT NOT NULL, client_type TEXT NOT NULL, layer TEXT NOT NULL, endpoint TEXT NOT NULL,
        requests BIGINT NOT NULL, successes BIGINT NOT NULL, failures BIGINT NOT NULL, submitted_units BIGINT NOT NULL,
        PRIMARY KEY(date,tenant_id,user_oid,host,client_type,layer,endpoint))`);
      await q('CREATE INDEX IF NOT EXISTS usage_tenant_date ON usage_analytics.api_usage_daily(tenant_id,date)');
      await q('CREATE INDEX IF NOT EXISTS usage_user_date ON usage_analytics.api_usage_daily(tenant_id,user_oid,date)');
      await q("INSERT INTO usage_analytics.meta VALUES ('schema_version','1'),('started_at',$1) ON CONFLICT DO NOTHING", [String(now)]);
    });
  } catch (error) { await client.end().catch(() => {}); throw error; }
  const ensureUser = (tenant, oid) => q('INSERT INTO usage_analytics.users(tenant_id,user_oid) VALUES ($1,$2) ON CONFLICT DO NOTHING', [tenant,oid]);
  async function record(r) {
    if (day(r.startedAt) < cutoff()) return;
    return transaction(async () => {
      await ensureUser(r.tenantId,r.userOid);
      await q(`INSERT INTO usage_analytics.api_usage_daily VALUES ($1,$2,$3,$4,$5,$6,$7,1,$8,$9,$10)
        ON CONFLICT(date,tenant_id,user_oid,host,client_type,layer,endpoint) DO UPDATE SET
        requests=api_usage_daily.requests+1, successes=api_usage_daily.successes+EXCLUDED.successes,
        failures=api_usage_daily.failures+EXCLUDED.failures, submitted_units=api_usage_daily.submitted_units+EXCLUDED.submitted_units`,
      [day(r.startedAt),r.tenantId,r.userOid,r.host,r.clientType,r.layer,r.endpoint,r.success ? 1 : 0,r.success ? 0 : 1,r.units]);
      if (r.layer === 'business' && r.endpoint !== 'detect') await q(`UPDATE usage_analytics.users SET
        first_used_at=LEAST(COALESCE(first_used_at,$1),$1), last_used_at=GREATEST(COALESCE(last_used_at,$1),$1)
        WHERE tenant_id=$2 AND user_oid=$3`, [r.startedAt,r.tenantId,r.userOid]);
    });
  }
  function where(f, translations = false) {
    const args = [f.from,f.to], terms = ['a.date BETWEEN $1 AND $2'];
    const bind = value => { args.push(value); return `$${args.length}`; };
    if (!f.tenants.includes('*')) terms.push(`a.tenant_id = ANY(${bind(f.tenants)}::text[])`);
    for (const [field,value] of [['host',f.host],['client_type',f.clientType],['user_oid',f.userOid]]) if (value) terms.push(`a.${field}=${bind(value)}`);
    if (f.search) { const b = bind(`%${f.search.replace(/[\\%_]/g, '\\$&')}%`); terms.push(`(u.display_name ILIKE ${b} ESCAPE '\\' OR u.mail ILIKE ${b} ESCAPE '\\')`); }
    if (translations) terms.push("a.layer='business' AND a.endpoint IN ('translate','translate-word')");
    return { args, sql: `FROM usage_analytics.api_usage_daily a LEFT JOIN usage_analytics.users u ON a.tenant_id=u.tenant_id AND a.user_oid=u.user_oid WHERE ${terms.join(' AND ')}` };
  }
  const numeric = new Set(['requests','successes','failures','submitted_units','users','days','business','upstream','first_used_at','last_used_at']);
  const normalize = rows => rows.map(row => Object.fromEntries(Object.entries(row).map(([key,value]) => [key, value !== null && numeric.has(key) ? Number(value) : value])));
  async function query(f, mode) {
    return transaction(async () => {
      const { sql,args } = where(f,mode === 'usage');
      const rows = async (select,suffix = '',extra = []) => normalize((await q(`${select} ${sql} ${suffix}`,[...args,...extra])).rows);
      const limit = `LIMIT $${args.length+1} OFFSET $${args.length+2}`, page = [f.pageSize,(f.page-1)*f.pageSize];
      const totals = 'SUM(a.requests) requests,SUM(a.successes) successes,SUM(a.failures) failures,SUM(a.submitted_units) submitted_units';
      const count = Number((await q(`SELECT COUNT(*) n FROM (SELECT a.tenant_id,a.user_oid ${sql} GROUP BY a.tenant_id,a.user_oid) people`,args)).rows[0].n);
      const coverage = { ...await meta(),retained_from:cutoff(),timezone:'Asia/Shanghai' };
      const common = { coverage,totalUsers:count,page:f.page,pageSize:f.pageSize };
      if (mode === 'usage') return { ...common,summary:{ ...(await rows('SELECT COALESCE(SUM(a.requests),0) requests'))[0],users:count },
        daily: await rows('SELECT a.date,SUM(a.requests) requests','GROUP BY a.date ORDER BY a.date'),
        activeDaily: Object.fromEntries((await q(`SELECT date,COUNT(*) n FROM (SELECT a.date,a.tenant_id,a.user_oid ${sql} GROUP BY a.date,a.tenant_id,a.user_oid) days GROUP BY date`,args)).rows.map(r => [r.date,Number(r.n)])),
        platforms: await rows('SELECT a.host,a.client_type,SUM(a.requests) requests,COUNT(DISTINCT (a.tenant_id,a.user_oid)) users','GROUP BY a.host,a.client_type ORDER BY a.host,a.client_type'),
        users: await rows(`SELECT a.tenant_id,a.user_oid,u.display_name,u.mail,u.first_used_at,u.last_used_at,
          COUNT(DISTINCT a.date) days,SUM(a.requests) requests,STRING_AGG(DISTINCT a.host || ' / ' || a.client_type,', ') platforms`,
        `GROUP BY a.tenant_id,a.user_oid,u.display_name,u.mail,u.first_used_at,u.last_used_at ORDER BY requests DESC,a.tenant_id,a.user_oid ${limit}`,page) };
      return { ...common, totals:await rows(`SELECT a.layer,${totals}`,'GROUP BY a.layer'),
        interfaces:await rows(`SELECT a.layer,a.endpoint,${totals}`,'GROUP BY a.layer,a.endpoint ORDER BY a.layer,a.endpoint'),
        daily:await rows(`SELECT a.date,a.layer,${totals}`,'GROUP BY a.date,a.layer ORDER BY a.date,a.layer'),
        platforms:await rows(`SELECT a.host,a.client_type,a.layer,${totals}`,'GROUP BY a.host,a.client_type,a.layer ORDER BY a.host,a.client_type,a.layer'),
        users:await rows(`SELECT a.tenant_id,a.user_oid,u.display_name,u.mail,
          SUM(CASE WHEN a.layer='business' THEN a.requests ELSE 0 END) business,
          SUM(CASE WHEN a.layer='upstream' THEN a.requests ELSE 0 END) upstream`,
        `GROUP BY a.tenant_id,a.user_oid,u.display_name,u.mail ORDER BY upstream DESC,business DESC,a.tenant_id,a.user_oid ${limit}`,page) };
    },true);
  }
  return { record,query,meta,
    gap: time => q("INSERT INTO usage_analytics.meta VALUES ('last_gap_at',$1) ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value",[String(time)]),
    profile: p => transaction(async () => { await ensureUser(p.tenantId,p.userOid); await q('UPDATE usage_analytics.users SET display_name=$1,mail=$2 WHERE tenant_id=$3 AND user_oid=$4',[p.displayName,p.mail,p.tenantId,p.userOid]); }),
    cleanup: time => q('DELETE FROM usage_analytics.api_usage_daily WHERE date < $1',[cutoff(time)]),
    deleteUser: (tenant,oid) => transaction(async () => {
      const usage = await q('DELETE FROM usage_analytics.api_usage_daily WHERE tenant_id=$1 AND user_oid=$2',[tenant,oid]);
      const users = await q('DELETE FROM usage_analytics.users WHERE tenant_id=$1 AND user_oid=$2',[tenant,oid]);
      return { usageRows:usage.rowCount,users:users.rowCount };
    }), close: () => client.end() };
}
module.exports = { openStore,day,cutoff };

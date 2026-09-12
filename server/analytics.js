const { AsyncLocalStorage } = require('async_hooks');
const { openStore } = require('./analytics-store');
const context = new AsyncLocalStorage();
const unavailable = () => new Error('Statistics unavailable');

function createAnalytics(connectionString, factory = openStore, retentionMonths) {
  let store, tail = Promise.resolve(), queued = 0, queries = 0, stopped = false, gapAt = 0, lastWarning = 0, retryAt = 0;
  function gap() {
    gapAt = Date.now();
    if (gapAt - lastWarning > 60000) { console.warn('Statistics unavailable or queue full; usage may be incomplete.'); lastWarning = gapAt; }
  }
  function send(method,value) {
    if (stopped || queued >= 500 || (method === 'query' && queries >= 4)) { if (method !== 'query') gap(); return Promise.reject(unavailable()); }
    queued++; if (method === 'query') queries++;
    const work = tail.then(async () => {
      if (!connectionString || Date.now() < retryAt) throw unavailable();
      try {
        store ||= await factory(connectionString, undefined, undefined, retentionMonths);
        if (gapAt) await store.gap(gapAt);
        if (method === 'query') return { ...await store.query(value.filters,value.mode),health:{lastGapAt:gapAt || null} };
        if (method === 'record') return await store.record(value);
        if (method === 'profile') return await store.profile(value);
        if (method === 'cleanup') return await store.cleanup(Date.now());
        if (method === 'adminSettings') return await store.adminSettings(value);
        if (method === 'saveAdminSettings') return await store.saveAdminSettings(value);
        return await store.meta();
      } catch (error) {
        if (error.status === 409) throw error;
        retryAt = Date.now()+5000;
        await store?.close().catch(() => {}); store = undefined;
        throw unavailable();
      }
    }).catch(error => { if (error.status !== 409) gap(); throw error; }).finally(() => { queued--; if (method === 'query') queries--; });
    tail = work.catch(() => {});
    if (method !== 'query') return work;
    return new Promise((resolve,reject) => {
      const timer = setTimeout(() => reject(unavailable()),5000);
      work.then(resolve,reject).finally(() => clearTimeout(timer));
    });
  }
  const ignore = promise => { promise.catch(() => {}); };
  if (connectionString) { ignore(send('ready')); ignore(send('cleanup')); }
  const cleanup = setInterval(() => ignore(send('cleanup')),86400000); cleanup.unref();
  return { record:value => ignore(send('record',value)),profile:value => ignore(send('profile',value)),
    query:(filters,mode) => send('query',{filters,mode}),flush:() => send('ready'),
    adminSettings:seed => send('adminSettings',seed),saveAdminSettings:value => send('saveAdminSettings',value),
    close:async () => { stopped=true; clearInterval(cleanup); await tail; await store?.close(); store=undefined; } };
}
function requestContext(identity,body,analytics) {
  return { tenantId:identity.tid.toLowerCase(),userOid:identity.oid.toLowerCase(),
    host:['word','outlook'].includes(body?.host) ? body.host : 'unknown',
    clientType:['online','local'].includes(body?.client_type) ? body.client_type : 'unknown',analytics };
}
function recordRequest(layer,endpoint,startedAt,success,units=0) {
  const current = context.getStore(); if (!current) return;
  const { analytics,...identity } = current;
  try { analytics.record({...identity,layer,endpoint,startedAt,success,units}); } catch { /* Never fail translation for statistics. */ }
}
module.exports = { createAnalytics,requestContext,recordRequest,context };

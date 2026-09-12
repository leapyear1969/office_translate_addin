const { createApp } = require('./app');
const { parseAdmins, filters } = require('./admin');
const tid = '11111111-1111-1111-1111-111111111111', oid = '22222222-2222-2222-2222-222222222222';
const admins = parseAdmins(JSON.stringify([{ tenant_id: tid, user_oid: oid, tenants: [tid] }]));
const { day } = require('./analytics-store');

test('anonymous and invalid-token floods cannot consume verified identity quotas', async () => {
  const app = createApp({origin:'https://test.invalid',analyticsAdmins:[
    ...admins,{tenant_id:tid,user_oid:tid,tenants:[tid]}, {tenant_id:oid,user_oid:oid,tenants:[oid]},
  ]},{analytics:{adminSettings:async entries=>({revision:1,entries})},authenticate:async token=>{
    if (token === 'invalid') throw new Error('invalid');
    return {tid:token === 'other-tenant' ? oid : tid,oid:token === 'other-user' ? tid : oid};
  }});
  expect(app.get('trust proxy')).toBe(false);
  const server = await new Promise(resolve=>{const s=app.listen(0,'127.0.0.1',()=>resolve(s));});
  const request = token=>fetch(`http://127.0.0.1:${server.address().port}/api/admin/session`, {
    headers:token ? {Authorization:`Bearer ${token}`} : {},
  });
  try {
    for(let i=0;i<120;i++) expect((await request(i%2 ? 'invalid' : '')).status).toBe(401);
    expect((await request('')).status).toBe(429);
    expect((await request('invalid')).status).toBe(429);
    for(let i=0;i<120;i++) expect((await request(i%2 ? 'owner-token-a' : 'owner-token-b')).status).toBe(200);
    expect((await request('owner-token-c')).status).toBe(429);
    expect((await request('other-user')).status).toBe(200);
    expect((await request('other-tenant')).status).toBe(200);
  } finally {server.closeAllConnections();await new Promise(resolve=>server.close(resolve));}
});

test('global grants require a configured identity and allow explicit tenant filtering', () => {
  const { allowedTenants } = require('./admin');
  for (const identity of [{ user_oid: oid }, { email: 'admin@example.invalid' }]) {
    expect(parseAdmins(JSON.stringify([{ tenant_id: tid, ...identity, tenants: ['*'] }]))[0].tenants).toEqual(['*']);
  }
  const globalAdmins = parseAdmins(JSON.stringify([{ tenant_id: tid, user_oid: oid, tenants: ['*'] }]));
  expect(allowedTenants({tid,oid},globalAdmins)).toEqual(['*']);
  expect(allowedTenants({tid,oid:tid},globalAdmins)).toEqual([]);
  expect(filters({}, ['*']).tenants).toEqual(['*']);
  expect(filters({tenant:oid}, ['*']).tenants).toEqual([oid]);
  expect(() => filters({tenant:'*'}, [tid])).toThrow();
  expect(() => filters({tenant:'invalid'}, ['*'])).toThrow();
});
test('configuration and filters fail closed', () => {
  expect(() => parseAdmins('[{"tenants":["*"]}]')).toThrow();
  expect(() => parseAdmins(JSON.stringify([{tenant_id:tid,user_oid:'bad',email:'admin@example.invalid',tenants:[tid]}]))).toThrow();
  for (const query of [{ from: '2026-02-30' }, { host: 'garbage' }, { tenant: oid }, { page: '1.1' }, { page_size: '201' }, { search: ['a','b'] }]) expect(() => filters(query, [tid])).toThrow();
  expect(filters({}, [tid])).toMatchObject({ tenants: [tid], to: day(Date.now()), page: 1, pageSize: 50 });
});

test('legacy email grants never authorize either original or recycled mailbox holders',async()=>{
  const emailAdmins=parseAdmins(JSON.stringify([{tenant_id:tid,email:'admin@example.invalid',tenants:[tid]}]));
  const profile=jest.fn(async(_token,identity)=>({id:identity.oid,tenantId:identity.tid,mail:'admin@example.invalid'}));
  const app=createApp({origin:'https://test.invalid',analyticsAdmins:emailAdmins},{
    analytics:{record(){},profile(){},query:async()=>({}),adminSettings:async seed=>({revision:1,entries:seed})},profile,
    authenticate:async token=>({tid,oid:token==='admin'?oid:tid,email:'admin@example.invalid'}),
  });
  const server=await new Promise(resolve=>{const s=app.listen(0,'127.0.0.1',()=>resolve(s));});
  const request=token=>fetch(`http://127.0.0.1:${server.address().port}/api/admin/session?email=admin@example.invalid`,{headers:{Authorization:`Bearer ${token}`}});
  try{
    expect((await request('spoof')).status).toBe(403);
    expect((await request('admin')).status).toBe(403);
    expect((await request('admin')).status).toBe(403);
    expect(profile).not.toHaveBeenCalled();
  }finally{server.closeAllConnections();await new Promise(resolve=>server.close(resolve));}
});
test('authenticated tenant-scoped admin routes, counters and profile whitelisting', async () => {
  const analytics = { record: jest.fn(), profile: jest.fn(), query: jest.fn(async () => ({ summary: { users: 1 } })),adminSettings:async seed=>({revision:1,entries:seed}) };
  const app = createApp({ origin: 'https://test.invalid', analyticsAdmins: admins }, {
    analytics, authenticate: async token => ({ tid, oid: token === 'admin' ? oid : tid }),
    profile: async () => ({ displayName: 'Name', mail: 'mail', photo: 'private-photo', token: 'secret' }),
    translate: async () => '<p>translated</p>', translateWord: async () => { throw new Error('private-text'); }, detect: async () => null,
  });
  const server = await new Promise(resolve => { const s = app.listen(0, '127.0.0.1', () => resolve(s)); });
  const request = (url, token = 'admin', body) => fetch(`http://127.0.0.1:${server.address().port}${url}`, { method: body ? 'POST' : 'GET', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
  try {
    expect((await request('/api/admin/overview', 'ordinary')).status).toBe(403);
    expect((await request(`/api/admin/overview?tenant=${oid}`)).status).toBe(403);
    expect((await request('/api/admin/overview')).status).toBe(200);
    expect(analytics.query).toHaveBeenCalledWith(expect.objectContaining({ tenants: [tid] }), 'usage');
    expect(analytics.record).not.toHaveBeenCalled();
    await request('/api/translate', 'admin', { html: 'secret source', to: 'en', host: 'word', client_type: 'online', tenant_id: 'spoof' });
    await request('/api/translate/word', 'admin', { paragraphs: ['secret'], to: 'en' });
    await request('/api/translate', 'admin', { html: '', to: 'en' });
    await request('/api/detect', 'admin', { html: 'secret' });
    expect(analytics.record.mock.calls.map(([r]) => [r.endpoint, r.success])).toEqual([['translate', true], ['translate-word', false], ['detect', true]]);
    expect(analytics.record.mock.calls[0][0]).toMatchObject({ tenantId: tid, userOid: oid, host: 'word', clientType: 'online' });
    expect(JSON.stringify(analytics.record.mock.calls)).not.toContain('secret');
    await request('/api/me');
    expect(analytics.profile).toHaveBeenCalledWith({ tenantId: tid, userOid: oid, displayName: 'Name', mail: 'mail', organizationName: null });
    expect((await request('/api/admin/delete', 'admin', {})).status).toBe(404);
    analytics.record.mockImplementation(() => { throw new Error('Database locked'); });
    expect((await request('/api/translate', 'admin', { html: 'text', to: 'en' })).status).toBe(200);
    // Admin reads use their own quota, not the 60/min business quota.
    for (let i = 0; i < 65; i++) expect((await request('/api/admin/session')).status).toBe(200);
    expect((await request('/api/detect', 'admin', { html: 'x' })).status).toBe(200);
  } finally { server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

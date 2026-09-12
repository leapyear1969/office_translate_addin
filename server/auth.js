const { jwtVerify, createRemoteJWKSet, decodeJwt } = require('jose');
const { ConfidentialClientApplication } = require('@azure/msal-node');

async function verifyIdentity(token, config, key, issuer) {
  const { payload } = await jwtVerify(token, key, {
    algorithms: ['RS256'], issuer, audience: [config.clientId, config.resource].filter(Boolean),
    requiredClaims: ['exp', 'iat', 'oid', 'tid', 'scp'],
  });
  if (payload.tid !== config.tenantId || typeof payload.oid !== 'string' || !payload.oid
    || !String(payload.scp).split(' ').includes('access_as_user')) throw new Error('Unauthorized identity');
  return payload;
}

function createAuth(config) {
  const metadata = new Map();
  const organizations = new Map();
  let cca;
  async function organizationName(accessToken, tid) {
    const key = tid.toLowerCase();
    const cached = organizations.get(key);
    if (cached && cached.expires > Date.now()) return cached.name;
    let name;
    try {
      const response = await fetch(`${config.graphBase}/v1.0/organization?$select=id,displayName`, {
        headers: { Authorization: `Bearer ${accessToken}` }, signal: AbortSignal.timeout(3000),
      });
      if (response.ok) {
        const data = await response.json();
        const organization = data.value?.find(o => typeof o.id === 'string' && o.id.toLowerCase() === key);
        if (typeof organization?.displayName === 'string' && organization.displayName.trim()) name = organization.displayName.trim().slice(0, 256);
      }
    } catch { /* Organization lookup is optional to login. */ }
    if (organizations.size >= 100) organizations.delete(organizations.keys().next().value);
    organizations.set(key, { name, expires: Date.now() + (name ? 86400000 : 300000) });
    return name;
  }
  async function discover(tenantId) {
    if (!config.clientId || !config.resource) {
      throw Object.assign(new Error('请先配置 SSO 的 Client ID 和 Application ID URI。'), { status: 503 });
    }
    if (!metadata.has(tenantId)) {
      if (metadata.size >= 100) metadata.delete(metadata.keys().next().value);
      const pending = Promise.all(['', '/v2.0'].map(async version => {
        const r = await fetch(`${config.authority}/${tenantId}${version}/.well-known/openid-configuration`, { signal: AbortSignal.timeout(10000) });
        if (!r.ok) throw new Error('Identity discovery unavailable');
        const doc = await r.json();
        return { issuer: doc.issuer, key: createRemoteJWKSet(new URL(doc.jwks_uri), { timeoutDuration: 10000 }) };
      })).catch(error => { if (metadata.get(tenantId) === pending) metadata.delete(tenantId); throw error; });
      metadata.set(tenantId, pending);
    }
    return metadata.get(tenantId);
  }
  async function authenticate(token) {
    // The unverified tid only selects a GUID path on the configured cloud.
    // Signature, issuer, audience, scope and this exact tid must all verify below.
    let tenantId;
    try { tenantId = decodeJwt(token).tid; } catch { /* Rejected below. */ }
    if (typeof tenantId !== 'string' || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(tenantId)) {
      throw Object.assign(new Error('SSO 令牌中的租户无效，请重新登录。'), { status: 401 });
    }
    const providers = await discover(tenantId);
    for (const provider of providers) {
      try { return await verifyIdentity(token, { ...config, tenantId }, provider.key, provider.issuer); }
      catch { /* Only metadata from the configured cloud is trusted. */ }
    }
    throw Object.assign(new Error('SSO 令牌无效或已过期，请重新登录。'), { status: 401 });
  }
  async function profile(token, identity) {
    if (!config.clientSecret) throw Object.assign(new Error('请在服务器 .env 中填写 CLIENT_SECRET 后重启服务。'), { status: 503 });
    cca ||= new ConfidentialClientApplication({ auth: {
      clientId: config.clientId, clientSecret: config.clientSecret,
      authority: `${config.authority}/organizations`,
    }, system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } } });
    try {
      const result = await cca.acquireTokenOnBehalfOf({ oboAssertion: token,
        authority: `${config.authority}/${identity.tid}`, scopes: [`${config.graphBase}/User.Read`] });
      const response = await fetch(`${config.graphBase}/v1.0/me?$select=id,displayName,mail,userPrincipalName`, {
        headers: { Authorization: `Bearer ${result.accessToken}` }, signal: AbortSignal.timeout(15000),
      });
      if (!response.ok) throw new Error('Graph profile failed');
      const user = await response.json();
      if (user.id !== identity.oid) throw new Error('Identity mismatch');
      // A missing or unavailable photo must not prevent sign-in.
      let photo;
      try {
        const image = await fetch(`${config.graphBase}/v1.0/me/photos/48x48/$value`, {
          headers: { Authorization: `Bearer ${result.accessToken}` }, signal: AbortSignal.timeout(3000),
        });
        if (image.ok && image.headers.get('content-type')?.split(';')[0] === 'image/jpeg') {
          const bytes = Buffer.from(await image.arrayBuffer());
          if (bytes.length <= 100000) photo = `data:image/jpeg;base64,${bytes.toString('base64')}`;
        }
      } catch { /* Use the display-name initial when a photo cannot be loaded. */ }
      const name = await organizationName(result.accessToken, identity.tid);
      return { ...(name ? { organizationName: name } : {}), ...(photo ? { photo } : {}), id: user.id, displayName: user.displayName || '', mail: user.mail || user.userPrincipalName || '', tenantId: identity.tid };
    } catch (error) {
      if (error.errorCode === 'consent_required' || error.subError === 'consent_required'
        || /AADSTS65001\b/.test(error.message || '')) {
        throw Object.assign(new Error('首次使用需要授权，请打开“翻译选项”，点击“登录并授权”。'), { status: 403, code: 'consent_required' });
      }
      throw Object.assign(new Error('读取用户信息失败，请检查客户端密钥、User.Read 权限及管理员同意。'), { status: 502 });
    }
  }
  return { authenticate, profile };
}
module.exports = { verifyIdentity, createAuth };

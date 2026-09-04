const { jwtVerify, createRemoteJWKSet } = require('jose');
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
  let metadata;
  let cca;
  async function discover() {
    if (!config.tenantId || !config.clientId || !config.resource) {
      throw Object.assign(new Error('请先配置 SSO 的 Tenant ID、Client ID 和 Application ID URI。'), { status: 503 });
    }
    if (!metadata) {
      metadata = Promise.all(['', '/v2.0'].map(async version => {
        const r = await fetch(`${config.authority}/${config.tenantId}${version}/.well-known/openid-configuration`, { signal: AbortSignal.timeout(10000) });
        if (!r.ok) throw new Error('Identity discovery unavailable');
        const doc = await r.json();
        return { issuer: doc.issuer, key: createRemoteJWKSet(new URL(doc.jwks_uri), { timeoutDuration: 10000 }) };
      })).catch(error => { metadata = undefined; throw error; });
    }
    return metadata;
  }
  async function authenticate(token) {
    const providers = await discover();
    for (const provider of providers) {
      try { return await verifyIdentity(token, config, provider.key, provider.issuer); }
      catch { /* Only trusted tenant metadata is tried; token claims never choose a host. */ }
    }
    throw Object.assign(new Error('SSO 令牌无效或已过期，请重新登录。'), { status: 401 });
  }
  async function profile(token, identity) {
    if (!config.clientSecret) throw Object.assign(new Error('请在服务器 .env 中填写 CLIENT_SECRET 后重启服务。'), { status: 503 });
    cca ||= new ConfidentialClientApplication({ auth: {
      clientId: config.clientId, clientSecret: config.clientSecret,
      authority: `${config.authority}/${config.tenantId}`,
    }, system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } } });
    try {
      const result = await cca.acquireTokenOnBehalfOf({ oboAssertion: token, scopes: [`${config.graphBase}/User.Read`] });
      const response = await fetch(`${config.graphBase}/v1.0/me?$select=id,displayName,mail,userPrincipalName`, {
        headers: { Authorization: `Bearer ${result.accessToken}` }, signal: AbortSignal.timeout(15000),
      });
      if (!response.ok) throw new Error('Graph profile failed');
      const user = await response.json();
      if (user.id !== identity.oid) throw new Error('Identity mismatch');
      return { id: user.id, displayName: user.displayName || '', mail: user.mail || user.userPrincipalName || '', tenantId: identity.tid };
    } catch {
      throw Object.assign(new Error('读取用户信息失败，请检查客户端密钥、User.Read 权限及管理员同意。'), { status: 502 });
    }
  }
  return { authenticate, profile };
}
module.exports = { verifyIdentity, createAuth };

const { randomBytes, createHash } = require('crypto');
const { ConfidentialClientApplication } = require('@azure/msal-node');

function createConsent(config) {
  const pending = new Map();
  const redirectUri = `${config.origin}/auth/consent/callback`;
  const random = () => randomBytes(32).toString('base64url');
  const scopes = [`${config.graphBase}/User.Read`];
  const unavailable = () => Object.assign(new Error('授权会话已过期，请关闭窗口后重新授权。'), { status: 400 });
  function take(state, remove = false) {
    const entry = typeof state === 'string' && pending.get(state);
    if (!entry || entry.expires < Date.now()) { pending.delete(state); throw unavailable(); }
    if (remove) pending.delete(state);
    return entry;
  }
  async function start(identity) {
    if (!config.clientSecret) throw Object.assign(new Error('请先在服务器配置 CLIENT_SECRET。'), { status: 503 });
    for (const [id, entry] of pending) if (entry.expires < Date.now()) pending.delete(id);
    if (pending.size >= 1000) throw Object.assign(new Error('授权请求繁忙，请稍后再试。'), { status: 503 });
    const state = random(), nonce = random(), verifier = random();
    const authority = `${config.authority}/${identity.tid}`;
    const client = new ConfidentialClientApplication({ auth: { clientId: config.clientId,
      clientSecret: config.clientSecret, authority },
    system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false } } });
    const url = await client.getAuthCodeUrl({ scopes, redirectUri, state, nonce, prompt: 'consent',
      codeChallenge: createHash('sha256').update(verifier).digest('base64url'), codeChallengeMethod: 'S256' });
    pending.set(state, { client, identity, verifier, nonce, url, launched: false, expires: Date.now() + 10 * 60000 });
    return { state, url: `${config.origin}/consent.html#state=${state}` };
  }
  function launch(state) {
    const entry = take(state);
    if (entry.launched) throw unavailable();
    entry.launched = true;
    return { url: entry.url };
  }
  async function complete(query) {
    const entry = take(query.state, true);
    if (!entry.launched) throw unavailable();
    let result = 'failed';
    if (query.error === 'access_denied') result = 'cancelled';
    else if (!query.error && typeof query.code === 'string') {
      try {
        const token = await entry.client.acquireTokenByCode({ code: query.code, scopes, redirectUri,
          codeVerifier: entry.verifier });
        const claims = token.idTokenClaims;
        if (claims?.tid === entry.identity.tid && claims?.oid === entry.identity.oid && claims?.nonce === entry.nonce) result = 'success';
        else result = 'account_mismatch';
      } catch { /* Never return tokens or raw identity provider errors to the dialog. */ }
    }
    return `${config.origin}/consent.html#state=${encodeURIComponent(query.state)}&result=${result}`;
  }
  return { start, launch, complete };
}
module.exports = { createConsent };

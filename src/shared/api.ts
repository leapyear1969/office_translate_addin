export interface UserProfile { id: string; displayName: string; mail: string; tenantId: string; photo?: string }
export interface Session { token: string; user: UserProfile }
declare const OfficeRuntime: { auth: { getAccessToken(options: { allowSignInPrompt: boolean; allowConsentPrompt: boolean }): Promise<string> } };

let pendingToken: Promise<string> | undefined;
let pendingSession: Promise<Session> | undefined;
let cachedSession: { session: Session; expiresAt: number } | undefined;
let sessionEpoch = 0;

// Reuse successful authentication briefly during live previews. Keep credentials
// in this runtime only; JWT claims here are a cache hint, never identity validation.
function sessionExpiresAt(token: string): number {
  try {
    const payload = token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
    const { exp } = JSON.parse(atob(payload.padEnd(Math.ceil(payload.length / 4) * 4, '=')));
    if (typeof exp === 'number' && Number.isFinite(exp)) {
      return Math.min(Date.now() + 300000, exp * 1000 - 60000);
    }
  } catch { /* Unknown token formats must not be cached. */ }
  return 0;
}

export function clearAuthentication(): void {
  cachedSession = undefined;
  ++sessionEpoch;
  // Do not unlock native SSO requests or reset a host-imposed cooldown.
}
let retryAfter = 0;
let cooldownMs = 60000;
function throttled() {
  return Object.assign(new Error(`Office SSO 请求过于频繁（13013），请在 ${Math.max(1, Math.ceil((retryAfter - Date.now()) / 1000))} 秒后重试。`),
    { code: '13013', retryAfter });
}
function getSsoToken(interactive: boolean): Promise<string> {
  if (Date.now() < retryAfter) return Promise.reject(throttled());
  if (!pendingToken) {
    // Keep the native request locked even when a caller's wait times out:
    // Office does not expose cancellation for getAccessToken.
    pendingToken = Promise.resolve().then(() => OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: interactive, allowConsentPrompt: interactive,
    })).then(token => { cooldownMs = 60000; return token; }).catch(error => {
      if (String(error?.code) === '13013') {
        retryAfter = Date.now() + cooldownMs;
        cooldownMs = Math.min(cooldownMs * 2, 300000);
        throw throttled();
      }
      throw error;
    }).finally(() => { pendingToken = undefined; });
  }
  return pendingToken;
}

export async function api<T>(path: string, token: string, body?: unknown): Promise<T> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 110000);
  try {
    const response = await fetch(path, { method: body === undefined ? 'GET' : 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body), signal: controller.signal,
    });
    if (response.status === 401 && cachedSession?.session.token === token) clearAuthentication();
    const result = await response.json();
    if (!response.ok) throw Object.assign(new Error(result.error || `请求失败（${response.status}）。`), { code: result.code });
    return result as T;
  } catch (error) {
    if (controller.signal.aborted) throw new Error('请求超时，原文保持不变，请重试。');
    throw error;
  } finally { clearTimeout(timer); }
}

export function authenticate(interactive = true): Promise<Session> {
  if (cachedSession && Date.now() < cachedSession.expiresAt) return Promise.resolve(cachedSession.session);
  cachedSession = undefined;
  // The first caller determines prompt permissions; automatic work must never
  // upgrade an in-flight silent request into an interactive request.
  if (!pendingSession) {
    const epoch = sessionEpoch;
    pendingSession = authenticateOnce(interactive).then(session => {
      if (epoch !== sessionEpoch) throw new Error('登录状态已变更，请重新登录。');
      cachedSession = { session, expiresAt: sessionExpiresAt(session.token) };
      return session;
    }).finally(() => { pendingSession = undefined; });
  }
  return pendingSession;
}

async function authenticateOnce(interactive: boolean): Promise<Session> {
  if (typeof OfficeRuntime === 'undefined' || !OfficeRuntime.auth?.getAccessToken) {
    throw new Error('当前环境不支持 Office SSO，请在支持 Office SSO 的 Office 客户端中打开插件。');
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  let token: string;
  try {
    token = await Promise.race([
      getSsoToken(interactive),
      new Promise<never>((_resolve, reject) => { timer = setTimeout(() => reject(new Error('SSO 登录超时，请重试。')), 20000); }),
    ]);
  } catch (error) {
    const value = error as { code?: number; message?: string };
    if (String(value.code) === '13013') throw error;
    if (String(value.code) === '13004') {
      throw new Error('SSO 登录失败（13004）：加载项清单的 SSO 资源地址无效。请检查页面地址与 WebApplicationInfo/Resource 的域名和端口是否一致，并确认 Resource 与 Entra Application ID URI 相同；修正后重新加载清单。');
    }
    throw new Error(value.code ? `SSO 登录失败（${value.code}），请检查应用预授权或在翻译选项中重试。` : value.message || 'SSO 登录失败。');
  } finally { clearTimeout(timer); }
  try { return { token, user: await api<UserProfile>('/api/me', token) }; }
  catch (error) {
    // Consent starts with the already verified Office identity, so retain the
    // assertion in memory for this error only; never put it in a URL or storage.
    if ((error as { code?: string }).code === 'consent_required') Object.assign(error as Error, { token });
    throw error;
  }
}

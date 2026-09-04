export interface UserProfile { id: string; displayName: string; mail: string; tenantId: string }
export interface Session { token: string; user: UserProfile }
declare const OfficeRuntime: { auth: { getAccessToken(options: { allowSignInPrompt: boolean; allowConsentPrompt: boolean }): Promise<string> } };

export async function api<T>(path: string, token: string, body?: unknown): Promise<T> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 110000);
  try {
    const response = await fetch(path, { method: body === undefined ? 'GET' : 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body), signal: controller.signal,
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || `请求失败（${response.status}）。`);
    return result as T;
  } catch (error) {
    if (controller.signal.aborted) throw new Error('请求超时，原文保持不变，请重试。');
    throw error;
  } finally { clearTimeout(timer); }
}

export async function authenticate(interactive = true): Promise<Session> {
  if (typeof OfficeRuntime === 'undefined' || !OfficeRuntime.auth?.getAccessToken) {
    throw new Error('当前环境不支持 Office SSO，请在 Outlook 中打开插件。');
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  let token: string;
  try {
    token = await Promise.race([
      OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: interactive, allowConsentPrompt: interactive }),
      new Promise<never>((_resolve, reject) => { timer = setTimeout(() => reject(new Error('SSO 登录超时，请重试。')), 20000); }),
    ]);
  } catch (error) {
    const value = error as { code?: number; message?: string };
    throw new Error(value.code ? `SSO 登录失败（${value.code}），请检查应用预授权或在翻译选项中重试。` : value.message || 'SSO 登录失败。');
  } finally { clearTimeout(timer); }
  return { token, user: await api<UserProfile>('/api/me', token) };
}

import { api } from './api';

export async function requestConsent(token: string): Promise<string> {
  const launch = await api<{ state: string; url: string }>('/api/consent/start', token, {});
  const url = new URL(launch.url);
  if (url.origin !== window.location.origin || url.protocol !== 'https:') throw new Error('授权窗口地址无效。');
  // No dialog messaging or token transfer is needed: the user reopens the add-in.
  // A visible link remains available when the host cannot open an external browser.
  try { Office.context.ui.openBrowserWindow(url.href); } catch { /* Use the link in the task pane. */ }
  return url.href;
}

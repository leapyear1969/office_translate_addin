// This page runs in an ordinary browser and must not wait for Office.onReady.
void (async () => {
  const params = new URLSearchParams(window.location.hash.slice(1));
  const state = params.get('state');
  const result = params.get('result');
  const status = document.getElementById('consent-status')!;
  if (!state || !/^[A-Za-z0-9_-]{43}$/.test(state)) { status.textContent = '授权会话无效，请关闭此窗口并重新打开插件重试。'; return; }
  if (result) {
    status.textContent = result === 'success' ? '授权已完成。请关闭此窗口，返回 Office，关闭并重新打开翻译插件。'
      : result === 'account_mismatch' ? '授权账户与 Office 账户不一致。请关闭此窗口并重新打开插件，使用当前 Office 账户授权。'
        : result === 'cancelled' ? '已取消授权。请关闭此窗口并重新打开插件重试。'
          : '授权未完成。请关闭此窗口并重新打开插件重试；如果微软提示需要管理员批准，请联系本租户管理员。';
    return;
  }
  try {
    const response = await fetch(`/auth/consent/launch/${encodeURIComponent(state)}`, { cache: 'no-store' });
    if (!response.ok) throw new Error('expired');
    const launch = await response.json();
    window.location.replace(launch.url);
  } catch { status.textContent = '无法启动授权或授权链接已过期。请关闭此窗口并重新打开插件重试。'; }
})();

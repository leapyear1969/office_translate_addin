/** @jest-environment jsdom */
test.each([
  ['success', '授权已完成。请关闭此窗口'],
  ['cancelled', '已取消授权'],
  ['account_mismatch', '授权账户与 Outlook 账户不一致'],
  ['failed', '授权未完成'],
])('shows browser completion state %s without Office runtime', (result, expected) => {
  jest.resetModules();
  delete (global as any).Office;
  document.body.innerHTML = '<p id="consent-status"></p>';
  window.location.hash = `state=${'a'.repeat(43)}&result=${result}`;
  require('./consent');
  expect(document.getElementById('consent-status')!.textContent).toContain(expected);
});

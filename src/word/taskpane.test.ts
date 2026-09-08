/** @jest-environment jsdom */
import { readFileSync } from 'fs';
import { join } from 'path';
const settle = async () => { for (let i = 0; i < 40; i++) await Promise.resolve(); };
const button = (id: string) => document.getElementById(id) as HTMLButtonElement;

async function setup(mode = 'never', excluded: string[] = []) {
  jest.resetModules();
  document.documentElement.innerHTML = readFileSync(join(__dirname, 'taskpane.html'), 'utf8');
  const translateDocument = jest.fn(async (..._args: any[]) => {});
  const restoreOriginalBody = jest.fn(async () => ({ restored: 1, skipped: 0 }));
  const api = jest.fn(async () => ({ language: 'en', score: 1 }));
  const authenticate = jest.fn(async () => ({ token: 'token', user: { displayName: 'User', mail: 'user@example.com' } }));
  jest.doMock('./document', () => ({ translateDocument, restoreOriginalBody }));
  jest.doMock('../shared/api', () => ({ api, authenticate }));
  const requestConsent = jest.fn(async () => 'https://example.com/auth');
  jest.doMock('../shared/consent', () => ({ requestConsent }));
  let saved: unknown = { mode, target: 'ja', excluded };
  let ready: Function = () => {};
  const commands: Record<string, Function> = {};
  const saveAsync = jest.fn((cb: Function) => cb({ status: 'succeeded' }));
  (globalThis as any).Office = {
    onReady: (cb: Function) => { ready = cb; }, HostType: { Word: 'Word' },
    AsyncResultStatus: { Succeeded: 'succeeded' }, VisibilityMode: { taskpane: 'Taskpane' },
    actions: { associate: (name: string, fn: Function) => { commands[name] = fn; } },
    addin: { showAsTaskpane: jest.fn(async () => {}), onVisibilityModeChanged: jest.fn(async () => {}) },
    context: { document: { settings: { get: () => saved, set: (_key: string, value: typeof saved) => { saved = value; }, saveAsync } } },
  };
  require('./taskpane'); await ready({ host: 'Word' }); await settle();
  return { translateDocument, restoreOriginalBody, commands, api, saveAsync, authenticate, requestConsent, savedSettings: () => saved };
}

test('manual scopes use saved language without detecting document language', async () => {
  const { translateDocument, api } = await setup();
  for (const scope of ['selection', 'paragraph', 'body']) {
    button(`translate-${scope}`).click(); await settle();
    expect(translateDocument).toHaveBeenLastCalledWith(scope, 'ja', expect.any(Function));
  }
  expect(api).not.toHaveBeenCalled();
  expect(document.body.textContent).not.toMatch(/Outlook|邮件|显示原文/);
});

test('repeated context commands translate to Chinese and always complete the event', async () => {
  const { commands, translateDocument } = await setup();
  const event = { completed: jest.fn() };
  await commands.translateSelectionChinese(event); await commands.translateSelectionChinese(event);
  expect(translateDocument).toHaveBeenCalledTimes(2);
  expect(translateDocument).toHaveBeenCalledWith('selection', 'zh-Hans', expect.any(Function));
  expect(event.completed).toHaveBeenCalledTimes(2);
});

test.each(['always', 'ask', 'never'])('legacy %s settings never detect, prompt or translate without a manual action', async mode => {
  const { translateDocument, api, savedSettings } = await setup(mode, ['en']);
  const visibility = (Office.addin.onVisibilityModeChanged as jest.Mock).mock.calls[0][0];
  visibility({ visibilityMode: 'Hidden' });
  visibility({ visibilityMode: 'Taskpane' });
  button('signin').click(); await settle();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en'; target.dispatchEvent(new Event('change', { bubbles: true }));
  button('preferences').dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  expect(savedSettings()).toEqual({ target: 'en' });
  expect(api).not.toHaveBeenCalled();
  expect(translateDocument).not.toHaveBeenCalled();
  expect(document.querySelector('input[name=mode]')).toBeNull();
  expect(document.getElementById('excluded-languages')).toBeNull();
  expect(document.getElementById('translation-prompt')).toBeNull();
  button('translate-selection').click(); await settle();
  expect(translateDocument).toHaveBeenCalledWith('selection', 'en', expect.any(Function));
});

test('failed settings save keeps the previous target and allows retry', async () => {
  const { saveAsync, translateDocument } = await setup();
  saveAsync.mockImplementationOnce(cb => cb({ status: 'failed' }));
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en'; target.dispatchEvent(new Event('change', { bubbles: true }));
  button('preferences').dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  expect(button('save-status').textContent).toContain('保存设置失败');
  button('translate-selection').click(); await settle();
  expect(translateDocument).toHaveBeenLastCalledWith('selection', 'ja', expect.any(Function));
  button('preferences').dispatchEvent(new Event('submit', { cancelable: true })); await settle();
  button('translate-selection').click(); await settle();
  expect(translateDocument).toHaveBeenLastCalledWith('selection', 'en', expect.any(Function));
});

test('explicit consent retains the browser authorization flow', async () => {
  const { authenticate, requestConsent } = await setup();
  authenticate.mockRejectedValueOnce(Object.assign(new Error('请授权'), { code: 'consent_required', token: 'sso' }));
  button('signin').click(); await settle();
  expect(requestConsent).toHaveBeenCalledWith('sso');
  expect(button('browser-consent').hidden).toBe(false);
});

test('restoration requires explicit confirmation and does not authenticate or translate again', async () => {
  const { restoreOriginalBody, translateDocument, authenticate } = await setup('always');
  translateDocument.mockClear(); authenticate.mockClear();
  button('restore-original').click(); await settle();
  expect(button('restore-prompt').hidden).toBe(false);
  expect(restoreOriginalBody).not.toHaveBeenCalled();
  button('cancel-restore').click(); await settle();
  expect(button('restore-prompt').hidden).toBe(true);
  expect(restoreOriginalBody).not.toHaveBeenCalled();
  button('restore-original').click(); button('confirm-restore').click(); await settle();
  expect(restoreOriginalBody).toHaveBeenCalledTimes(1);
  expect(authenticate).not.toHaveBeenCalled();
  expect(translateDocument).not.toHaveBeenCalled();
  expect(button('status').textContent).toContain('已恢复 1 处翻译的原文');
});

test('failed restoration reports an error and permits retry', async () => {
  const { restoreOriginalBody } = await setup();
  restoreOriginalBody.mockRejectedValueOnce(new Error('恢复失败'));
  button('restore-original').click(); button('confirm-restore').click(); await settle();
  expect(button('status').textContent).toBe('恢复失败');
  expect(button('restore-original').disabled).toBe(false);
  button('restore-original').click(); button('confirm-restore').click(); await settle();
  expect(restoreOriginalBody).toHaveBeenCalledTimes(2);
});

test('logout blocks document commands and visibility login until explicit login', async () => {
  const { authenticate, translateDocument, commands } = await setup();
  button('account-avatar').click(); button('signout').click();
  expect(button('account-name').textContent).toBe('已注销');
  expect(button('translate-selection').disabled).toBe(true);
  const visibility = (Office.addin.onVisibilityModeChanged as jest.Mock).mock.calls[0][0];
  visibility({ visibilityMode: 'Taskpane' }); await settle();
  await commands.translateSelectionChinese({ completed: jest.fn() });
  expect(translateDocument).not.toHaveBeenCalled();
  expect(authenticate).toHaveBeenCalledTimes(1);
  button('account-avatar').click(); button('signin').click(); await settle();
  expect(authenticate).toHaveBeenLastCalledWith(true);
  expect(button('translate-selection').disabled).toBe(false);
});

test('logout invalidates a pending login response', async () => {
  const { authenticate } = await setup();
  let resolve: Function = () => {};
  authenticate.mockImplementationOnce(() => new Promise<any>(done => { resolve = done; }));
  button('signin').click();
  button('signout').click();
  resolve({ token: 'late', user: { displayName: 'Late', mail: 'late@example.com' } });
  await settle();
  expect(button('account-name').textContent).toBe('已注销');
  expect(button('translate-selection').disabled).toBe(true);
});


test('restoration reports skipped edits in the panel', async () => {
  const { restoreOriginalBody } = await setup();
  restoreOriginalBody.mockResolvedValueOnce({ restored: 1, skipped: 2 });
  button('restore-original').click(); button('confirm-restore').click(); await settle();
  expect(button('status').textContent).toContain('已恢复 1 处翻译；2 处');
  expect(button('status').textContent).toContain('已保留当前编辑');
});

test('full-body completion reports paragraphs retained for safety', async () => {
  const { translateDocument } = await setup();
  translateDocument.mockResolvedValueOnce({ htmlBackup: false, skippedParagraphs: 2 } as any);
  button('translate-body').click(); await settle();
  expect(button('status').textContent).toContain('2 个段落');
  expect(button('status').textContent).toContain('已保留原文');
});

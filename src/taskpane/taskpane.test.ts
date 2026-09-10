/** @jest-environment jsdom */
import { readFileSync } from 'fs';
import { join } from 'path';

const settle = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };
const button = (id: string) => document.getElementById(id) as HTMLButtonElement;

test('places a decorative translation icon before the options title', () => {
  document.documentElement.innerHTML = readFileSync(join(__dirname, 'taskpane.html'), 'utf8');
  const heading = document.querySelector('h1')!;
  const icon = heading.firstElementChild as HTMLImageElement;
  expect(icon?.tagName).toBe('IMG');
  expect(icon.getAttribute('src')).toBe('assets/icon-32.png');
  expect(icon.getAttribute('alt')).toBe('');
  expect(heading.textContent?.trim()).toBe('翻译选项');
});

test('places the initially hidden translation prompt before the options header', () => {
  document.documentElement.innerHTML = readFileSync(join(__dirname, 'taskpane.html'), 'utf8');
  const prompt = document.getElementById('translation-prompt')!;
  expect(document.querySelector('main')!.firstElementChild).toBe(prompt);
  expect(prompt.hidden).toBe(true);
  expect(prompt.querySelector('#translate-now')).not.toBeNull();
  expect(prompt.querySelector('#dismiss')).not.toBeNull();
});

test('removes the standalone original and translation buttons without breaking initialization', async () => {
  await setup();
  expect(document.getElementById('show-original')).toBeNull();
  expect(document.getElementById('translate-message')).toBeNull();
  expect(document.querySelector('.message-actions')).toBeNull();
  expect((document.getElementById('settings-fields') as HTMLFieldSetElement).disabled).toBe(false);
});

test('translates from the top prompt and hides it on completion', async () => {
  const { api, setAsync } = await setup(undefined, 'ask', 'ja');
  expect(button('translation-prompt').hidden).toBe(false);
  expect(button('prompt-message').textContent).toBe('这封邮件使用英语，是否翻译为日语？');
  button('translate-now').click();
  await settle();
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'ja' });
  expect(setAsync).toHaveBeenCalledWith('<p>译文</p>', { coercionType: 'html' }, expect.any(Function));
  expect(button('translation-prompt').hidden).toBe(true);
  expect(button('status').textContent).toBe('翻译完成。点击提示栏中的“显示原文”或重新打开邮件即可查看原文。');
});

test('dismisses the top prompt without translating the message', async () => {
  const { api, setAsync } = await setup(undefined, 'ask');
  expect(button('translation-prompt').hidden).toBe(false);
  button('dismiss').click();
  expect(button('translation-prompt').hidden).toBe(true);
  expect(api).toHaveBeenCalledTimes(1);
  expect(api).toHaveBeenCalledWith('/api/detect', 'token', { html: '<p>Original</p>' });
  expect(setAsync).not.toHaveBeenCalled();
});

test('opens the header account popup and closes it with Escape or an outside click', async () => {
  await setup();
  expect(document.querySelector('.pane-header #account-avatar')).not.toBeNull();
  expect(button('account-menu').hidden).toBe(true);
  button('account-avatar').click();
  expect(button('account-menu').hidden).toBe(false);
  expect(document.activeElement).toBe(button('signin'));
  expect(button('account-name').textContent).toBe('User');
  expect(button('account-email').textContent).toBe('u@example.com');
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
  expect(button('account-menu').hidden).toBe(true);
  expect(document.activeElement).toBe(button('account-avatar'));
  button('account-avatar').click();
  document.body.click();
  expect(button('account-menu').hidden).toBe(true);
});

test('logout blocks translation and silent login on mail changes until explicit login', async () => {
  const { authenticate, handlers } = await setup();
  button('account-avatar').click(); button('signout').click();
  expect(button('account-name').textContent).toBe('已注销');
  expect(button('translate-now').disabled).toBe(true);
  handlers.itemChanged(); await settle();
  expect(authenticate).toHaveBeenCalledTimes(1);
  button('account-avatar').click(); button('signin').click(); await settle();
  expect(authenticate).toHaveBeenLastCalledWith(true);
  expect(button('translate-now').disabled).toBe(false);
});

async function setup(context?: unknown, mode = 'never', language = 'zh-Hans', deferInitialization = false) {
  jest.resetModules();
  document.documentElement.innerHTML = readFileSync(join(__dirname, 'taskpane.html'), 'utf8');
  const api = jest.fn(async (path: string) => path === '/api/detect' ? { language: 'en', score: 1 } : { html: '<p>译文</p>' });
  const authenticate = jest.fn(async () => ({ token: 'token', user: { id: 'u', displayName: 'User', mail: 'u@example.com' } }));
  jest.doMock('../shared/api', () => ({ api, authenticate, clearAuthentication: jest.fn() }));
  const requestConsent = jest.fn(async () => 'https://addin.example/consent.html#state=test');
  jest.doMock('../shared/consent', () => ({ requestConsent }));
  const handlers: Record<string, Function> = {};
  const itemHandlers: Record<string, Function> = {};
  const notifications = new Map<string, Office.NotificationMessageDetails>();
  const setAsync = jest.fn((_html, _options, cb) => cb({ status: 'succeeded' }));
  const item = {
    itemId: 'a',
    body: { getAsync: jest.fn((_type, cb) => cb({ status: 'succeeded', value: '<p>Original</p>' })) },
    display: { body: { setAsync } },
    notificationMessages: {
      replaceAsync: jest.fn((key, value, cb) => { notifications.set(key, { ...value, key }); cb({ status: 'succeeded' }); }),
      getAllAsync: jest.fn(cb => cb({ status: 'succeeded', value: [...notifications.values()] })),
    },
    getInitializationContextAsync: jest.fn(cb => { if (!deferInitialization) cb({ status: 'succeeded', value: context }); }),
    addHandlerAsync: jest.fn((type, handler, cb) => { itemHandlers[type] = handler; cb?.({ status: 'succeeded' }); }),
  };
  let saved = { mode, target: language, excluded: [] };
  let ready: Function = () => {};
  (globalThis as any).Office = {
    onReady: (cb: Function) => { ready = cb; }, HostType: { Outlook: 'Outlook' },
    AsyncResultStatus: { Succeeded: 'succeeded' }, CoercionType: { Html: 'html' },
    EventType: { ItemChanged: 'itemChanged', InitializationContextChanged: 'contextChanged' },
    MailboxEnums: { ItemNotificationMessageType: { InformationalMessage: 'info', InsightMessage: 'insight' } },
    context: {
      mailbox: { item, addHandlerAsync: (type: string, handler: Function, cb: Function) => { handlers[type] = handler; cb({ status: 'succeeded' }); } },
      roamingSettings: { get: () => saved, set: (_key: string, value: typeof saved) => { saved = value; }, saveAsync: (cb: Function) => cb({ status: 'succeeded' }) },
    },
  };
  require('./taskpane');
  ready({ host: 'Outlook' });
  await settle();
  return { api, authenticate, requestConsent, item, setAsync, itemHandlers, handlers, notifications };
}

test('user consent opens a browser and asks to reopen the add-in without premature sign-in retry', async () => {
  const { authenticate, requestConsent } = await setup();
  authenticate.mockRejectedValueOnce(Object.assign(new Error('请授权'), { code: 'consent_required', token: 'sso' }));
  button('signin').click();
  await settle();
  expect(requestConsent).toHaveBeenCalledWith('sso');
  expect(document.getElementById('account-name')!.textContent).toBe('等待浏览器授权');
  expect(document.getElementById('browser-consent')!.hidden).toBe(false);
  expect((document.getElementById('browser-consent-link') as HTMLAnchorElement).href).toBe('https://addin.example/consent.html#state=test');
  expect(authenticate).toHaveBeenCalledTimes(2);
  expect(button('signin').disabled).toBe(false);
});

test('silent missing consent offers a button without opening a dialog and cancellation allows retry', async () => {
  const { authenticate, requestConsent, handlers } = await setup();
  const error = Object.assign(new Error('请授权'), { code: 'consent_required', token: 'sso' });
  authenticate.mockRejectedValue(error);
  handlers.itemChanged();
  await settle();
  expect(requestConsent).not.toHaveBeenCalled();
  expect(button('signin').textContent).toBe('登录并授权');
  requestConsent.mockRejectedValueOnce(new Error('已取消授权'));
  button('signin').click();
  await settle();
  expect(button('signin').disabled).toBe(false);
  expect(document.getElementById('status')!.textContent).toBe('已取消授权');
});

test('translates from a notification with the saved target language even in never mode', async () => {
  const { api, setAsync, itemHandlers } = await setup(undefined, 'never', 'ja');
  expect(button('translation-prompt').hidden).toBe(true);
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'ja' });
  expect(setAsync).toHaveBeenCalledWith('<p>译文</p>', { coercionType: 'html' }, expect.any(Function));
});

test('updates the notification translation target only after saving preferences', async () => {
  const { api, itemHandlers } = await setup();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  expect(api).toHaveBeenLastCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'zh-Hans' });
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  expect(api).toHaveBeenLastCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'en' });
});

test.each([JSON.stringify({ action: 'showOriginal' }), { action: 'showOriginal' }])('restores on notification launch without auto-translating it again: %p', async context => {
  const { setAsync, api } = await setup(context, 'always');
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(setAsync).toHaveBeenCalledWith('<p>Original</p>', { coercionType: 'html' }, expect.any(Function));
  expect(api).not.toHaveBeenCalled();
  expect(document.getElementById('status')!.textContent).toBe('已显示原文。');
});

test.each([JSON.stringify({ action: 'translateMessage' }), { action: 'translateMessage' }])('translates once on notification launch using saved preferences even in never mode: %p', async context => {
  const { api, setAsync, item } = await setup(context, 'never', 'ja');
  expect(api).toHaveBeenCalledTimes(1);
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'ja' });
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(item.notificationMessages.replaceAsync).toHaveBeenLastCalledWith('mail-translation-status',
    expect.objectContaining({ actions: [expect.objectContaining({ actionText: '显示原文' })] }), expect.any(Function));
});

test('does not automatically inspect or translate again after an explicit translation launch', async () => {
  const { api, setAsync } = await setup({ action: 'translateMessage' }, 'always');
  expect(api).toHaveBeenCalledTimes(1);
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'zh-Hans' });
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(button('status').textContent).toContain('翻译完成');
});

test('cycles between original and translation using the actual notification contexts in an open pane', async () => {
  const { api, item, itemHandlers, setAsync } = await setup({ action: 'showOriginal' }, 'always');
  const notification = () => item.notificationMessages.replaceAsync.mock.calls.slice(-1)[0][1];
  expect(notification().actions).toEqual([expect.objectContaining({ actionText: '将邮件翻译为：中文（简体）' })]);
  itemHandlers.contextChanged({ initializationContextData: notification().actions[0].contextData });
  await settle();
  expect(notification().actions).toEqual([expect.objectContaining({ actionText: '显示原文' })]);
  itemHandlers.contextChanged({ initializationContextData: notification().actions[0].contextData });
  await settle();
  expect(notification().actions).toEqual([expect.objectContaining({ actionText: '将邮件翻译为：中文（简体）' })]);
  expect(setAsync.mock.calls.map(call => call[0])).toEqual(['<p>Original</p>', '<p>译文</p>', '<p>Original</p>']);
  expect(api).toHaveBeenCalledTimes(1);
  expect(button('status').textContent).toBe('已显示原文。');
});

test('refreshes the restored-message notification target only after saving preferences', async () => {
  const { api, item, itemHandlers } = await setup({ action: 'showOriginal' });
  const notification = () => item.notificationMessages.replaceAsync.mock.calls.slice(-1)[0][1];
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  expect(notification().actions).toEqual([expect.objectContaining({ actionText: '将邮件翻译为：中文（简体）' })]);
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  expect(notification().actions).toEqual([expect.objectContaining({ actionText: '将邮件翻译为：英语' })]);
  expect(api).not.toHaveBeenCalled();
  itemHandlers.contextChanged({ initializationContextData: notification().actions[0].contextData });
  await settle();
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'en' });
});

test('saving preferences after a failed restore does not claim the original is displayed', async () => {
  const { item, setAsync, itemHandlers } = await setup();
  setAsync.mockImplementation((_html, _options, cb) => cb({ status: 'failed' }));
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  expect(item.notificationMessages.replaceAsync).not.toHaveBeenCalled();
});

test('does not change a translated notification to original when preferences are saved', async () => {
  const { item, itemHandlers } = await setup({ action: 'showOriginal' });
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  item.notificationMessages.replaceAsync.mockClear();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  expect(item.notificationMessages.replaceAsync).not.toHaveBeenCalled();
});

test('saving preferences preserves a show-original notification created by the ribbon command', async () => {
  const { item } = await setup({ action: 'showOriginal' });
  const { translateCurrentMessage } = require('../shared/mail');
  await translateCurrentMessage('zh-Hans');
  item.notificationMessages.replaceAsync.mockClear();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  expect(item.notificationMessages.replaceAsync).not.toHaveBeenCalled();
});

test('saving preferences does not recreate a notification the user dismissed', async () => {
  const { item, notifications } = await setup({ action: 'showOriginal' });
  notifications.clear();
  item.notificationMessages.replaceAsync.mockClear();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  expect(item.notificationMessages.replaceAsync).not.toHaveBeenCalled();
});

test('restoring during a failed settings save uses the last committed target', async () => {
  const { item, itemHandlers, api } = await setup();
  let finishSave!: (result: any) => void;
  (Office.context.roamingSettings as any).saveAsync = (cb: Function) => { finishSave = cb as typeof finishSave; };
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  finishSave({ status: 'failed' });
  await settle();
  const notification = item.notificationMessages.replaceAsync.mock.calls.slice(-1)[0][1];
  expect(notification.actions).toEqual([expect.objectContaining({ actionText: '将邮件翻译为：中文（简体）' })]);
  expect(button('save-status').textContent).toContain('保存设置失败');
  itemHandlers.contextChanged({ initializationContextData: notification.actions[0].contextData });
  await settle();
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'zh-Hans' });
});

test('a settings save completing during restore is reflected in the final notification', async () => {
  const { item, setAsync, itemHandlers } = await setup();
  let finishRestore!: (result: any) => void;
  setAsync.mockImplementationOnce((_html, _options, cb) => { finishRestore = cb; });
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  finishRestore({ status: 'succeeded' });
  await settle();
  expect(item.notificationMessages.replaceAsync).toHaveBeenLastCalledWith('mail-translation-status',
    expect.objectContaining({ actions: [expect.objectContaining({ actionText: '将邮件翻译为：英语' })] }), expect.any(Function));
});

test.each([undefined, { action: 'translateMessage' }])('does not replay stale launch data or auto-translate after a newer notification action: %p', async context => {
  const { item, itemHandlers, api, setAsync } = await setup(undefined, 'always', 'zh-Hans', true);
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  item.getInitializationContextAsync.mock.calls[0][0]({ status: 'succeeded', value: context });
  await settle();
  expect(api).toHaveBeenCalledTimes(1);
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'zh-Hans' });
  expect(setAsync).toHaveBeenCalledTimes(1);
});

test('a late initial context read does not supersede sign-in for a newly selected item', async () => {
  const { item, handlers, api, authenticate, setAsync } = await setup(undefined, 'always', 'zh-Hans', true);
  let finishSignIn!: (value: any) => void;
  authenticate.mockImplementationOnce(() => new Promise(resolve => { finishSignIn = resolve; }));
  (Office.context.mailbox as any).item = { ...item, itemId: 'b' };
  handlers.itemChanged();
  await settle();
  item.getInitializationContextAsync.mock.calls[0][0]({ status: 'succeeded', value: undefined });
  await settle();
  finishSignIn({ token: 'token', user: { id: 'u', displayName: 'User', mail: 'u@example.com' } });
  await settle();
  expect(authenticate).toHaveBeenCalledTimes(2);
  expect(api).toHaveBeenCalledTimes(2);
  expect(setAsync).toHaveBeenCalledTimes(1);
});

test('ignores repeated notification translation clicks while translation is pending', async () => {
  const { api, itemHandlers, setAsync } = await setup();
  let finishTranslation!: (value: any) => void;
  api.mockImplementationOnce(() => new Promise(resolve => { finishTranslation = resolve; }));
  const event = { initializationContextData: { action: 'translateMessage' } };
  itemHandlers.contextChanged(event);
  await settle();
  itemHandlers.contextChanged(event);
  await settle();
  expect(api).toHaveBeenCalledTimes(1);
  expect(button('translate-now').disabled).toBe(true);
  finishTranslation({ html: '<p>译文</p>' });
  await settle();
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(button('translate-now').disabled).toBe(false);
});

test('does not automatically inspect after an in-flight sign-in is superseded by a notification action', async () => {
  const { api, authenticate, handlers, itemHandlers, setAsync } = await setup(undefined, 'always');
  api.mockClear(); setAsync.mockClear();
  let finishSignIn!: (value: any) => void;
  authenticate.mockImplementationOnce(() => new Promise(resolve => { finishSignIn = resolve; }));
  handlers.itemChanged();
  await settle();
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  finishSignIn({ token: 'token', user: { id: 'u', displayName: 'User', mail: 'u@example.com' } });
  await settle();
  expect(api).toHaveBeenCalledTimes(1);
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'zh-Hans' });
  expect(setAsync).toHaveBeenCalledTimes(1);
});

test('ignores a translate notification from a previously selected message', async () => {
  const { api, handlers, itemHandlers, item, setAsync } = await setup();
  const previousHandler = itemHandlers.contextChanged;
  (Office.context.mailbox as any).item = { ...item, itemId: 'b' };
  handlers.itemChanged();
  await settle();
  previousHandler({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  expect(api).not.toHaveBeenCalled();
  expect(setAsync).not.toHaveBeenCalled();
});

test('notification translation does not bypass explicit sign-out', async () => {
  const { api, authenticate, itemHandlers, setAsync } = await setup();
  button('account-avatar').click(); button('signout').click();
  authenticate.mockClear();
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  expect(authenticate).not.toHaveBeenCalled();
  expect(api).not.toHaveBeenCalled();
  expect(setAsync).not.toHaveBeenCalled();
});

test('normal options launch does not restore the body', async () => {
  const { setAsync } = await setup();
  expect(setAsync).not.toHaveBeenCalled();
});

test('handles a notification while the options pane is already open', async () => {
  const { itemHandlers, setAsync } = await setup();
  expect(itemHandlers.contextChanged).toEqual(expect.any(Function));
  itemHandlers.contextChanged({ initializationContextData: JSON.stringify({ action: 'showOriginal' }) });
  await settle();
  expect(setAsync).toHaveBeenCalledWith('<p>Original</p>', { coercionType: 'html' }, expect.any(Function));
});

test('notification restore does not need authentication and reports failure in the options pane', async () => {
  const { authenticate, setAsync, itemHandlers } = await setup();
  authenticate.mockClear();
  setAsync.mockImplementationOnce((_html, _options, cb) => cb({ status: 'failed' }));
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  expect(authenticate).not.toHaveBeenCalled();
  expect(document.getElementById('status')!.textContent).toContain('原文显示失败');
  expect(button('translate-now').disabled).toBe(false);
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  expect(setAsync).toHaveBeenCalledTimes(2);
  expect(button('status').textContent).toBe('已显示原文。');
});

test('ignores unknown initialization data', async () => {
  const { setAsync } = await setup('not JSON');
  expect(setAsync).not.toHaveBeenCalled();
});

test('notification restoration invalidates a pending automatic language detection', async () => {
  const { api, setAsync, handlers, itemHandlers } = await setup(undefined, 'always');
  api.mockClear(); setAsync.mockClear();
  let finishDetection!: (value: any) => void;
  api.mockImplementationOnce(() => new Promise(resolve => { finishDetection = resolve; }));
  handlers.itemChanged();
  await settle();
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  finishDetection({ language: 'en', score: 1 });
  await settle();
  expect(api).toHaveBeenCalledTimes(1);
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(setAsync).toHaveBeenCalledWith('<p>Original</p>', { coercionType: 'html' }, expect.any(Function));
});

test('disables the prompt action during translation and reenables it afterward', async () => {
  const { api, setAsync, itemHandlers } = await setup(undefined, 'ask');
  let finishTranslation!: (value: any) => void;
  api.mockImplementationOnce(() => new Promise(resolve => { finishTranslation = resolve; }));
  button('translate-now').click();
  await settle();
  expect(button('translate-now').disabled).toBe(true);
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  expect(setAsync).not.toHaveBeenCalled();
  finishTranslation({ html: '<p>译文</p>' });
  await settle();
  expect(button('translate-now').disabled).toBe(false);
  expect(setAsync).toHaveBeenCalledTimes(1);
});

test('resumes automatic translation after switching away from a restored message', async () => {
  const { handlers, setAsync, api, item } = await setup({ action: 'showOriginal' }, 'always');
  setAsync.mockClear();
  (Office.context.mailbox as any).item = { ...item, itemId: 'b' };
  handlers.itemChanged();
  await settle();
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'zh-Hans' });
  expect(setAsync).toHaveBeenCalledWith('<p>译文</p>', { coercionType: 'html' }, expect.any(Function));
});

test('disables mail actions when there is no current message', async () => {
  const { handlers } = await setup();
  (Office.context.mailbox as any).item = null;
  handlers.itemChanged();
  await settle();
  expect(button('translate-now').disabled).toBe(true);
  expect(button('translation-prompt').hidden).toBe(true);
});

test('can restore the new message while a previous message translation is still pending', async () => {
  const { api, handlers, itemHandlers, item, setAsync } = await setup();
  let finishTranslation!: (value: any) => void;
  api.mockImplementationOnce(() => new Promise(resolve => { finishTranslation = resolve; }));
  itemHandlers.contextChanged({ initializationContextData: { action: 'translateMessage' } });
  await settle();
  (Office.context.mailbox as any).item = { ...item, itemId: 'b' };
  handlers.itemChanged();
  await settle();
  itemHandlers.contextChanged({ initializationContextData: { action: 'showOriginal' } });
  await settle();
  expect(setAsync).toHaveBeenCalledWith('<p>Original</p>', { coercionType: 'html' }, expect.any(Function));
  finishTranslation({ html: '<p>旧邮件译文</p>' });
  await settle();
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(document.getElementById('status')!.textContent).toBe('已显示原文。');
});

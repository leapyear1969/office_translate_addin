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

test('shows compact account details and sign-in before the message actions', async () => {
  await setup();
  const account = document.querySelector('.account')!;
  expect(account.nextElementSibling?.className).toBe('message-actions');
  expect(account.querySelector('.account-details #account-name')?.textContent).toBe('User');
  expect(account.querySelector('.account-details #account-email')?.textContent).toBe('u@example.com');
  expect(account.querySelector('.account-label')).toBeNull();
  expect(button('signin').parentElement).toBe(account);
  expect(button('signin').hidden).toBe(false);
});

async function setup(context?: unknown, mode = 'never', language = 'zh-Hans') {
  jest.resetModules();
  document.documentElement.innerHTML = readFileSync(join(__dirname, 'taskpane.html'), 'utf8');
  const api = jest.fn(async (path: string) => path === '/api/detect' ? { language: 'en', score: 1 } : { html: '<p>译文</p>' });
  const authenticate = jest.fn(async () => ({ token: 'token', user: { id: 'u', displayName: 'User', mail: 'u@example.com' } }));
  jest.doMock('../shared/api', () => ({ api, authenticate }));
  const handlers: Record<string, Function> = {};
  const itemHandlers: Record<string, Function> = {};
  const setAsync = jest.fn((_html, _options, cb) => cb({ status: 'succeeded' }));
  const item = {
    itemId: 'a',
    body: { getAsync: jest.fn((_type, cb) => cb({ status: 'succeeded', value: '<p>Original</p>' })) },
    display: { body: { setAsync } },
    notificationMessages: { replaceAsync: jest.fn((_key, _value, cb) => cb({ status: 'succeeded' })) },
    getInitializationContextAsync: jest.fn(cb => cb({ status: 'succeeded', value: context })),
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
  return { api, authenticate, item, setAsync, itemHandlers, handlers };
}

test('shows both actions with the saved target language and translates even in never mode', async () => {
  const { api, setAsync } = await setup(undefined, 'never', 'ja');
  expect(button('show-original')).not.toBeNull();
  expect(button('translate-message')?.textContent).toBe('将邮件翻译为：日语');
  button('translate-message').click();
  await settle();
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'ja' });
  expect(setAsync).toHaveBeenCalledWith('<p>译文</p>', { coercionType: 'html' }, expect.any(Function));
});

test('updates the action label and actual target only after saving preferences', async () => {
  const { api } = await setup();
  const target = document.getElementById('target-language') as HTMLSelectElement;
  target.value = 'en';
  target.dispatchEvent(new Event('change', { bubbles: true }));
  expect(button('translate-message')?.textContent).toBe('将邮件翻译为：中文（简体）');
  document.getElementById('preferences')!.dispatchEvent(new Event('submit', { cancelable: true }));
  await settle();
  expect(button('translate-message').textContent).toBe('将邮件翻译为：英语');
  button('translate-message').click();
  await settle();
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'en' });
});

test.each([JSON.stringify({ action: 'showOriginal' }), { action: 'showOriginal' }])('restores on notification launch without auto-translating it again: %p', async context => {
  const { setAsync, api } = await setup(context, 'always');
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(setAsync).toHaveBeenCalledWith('<p>Original</p>', { coercionType: 'html' }, expect.any(Function));
  expect(api).not.toHaveBeenCalled();
  expect(document.getElementById('status')!.textContent).toBe('已显示原文。');
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

test('manual restore does not need authentication and reports failure in the options pane', async () => {
  const { authenticate, setAsync } = await setup();
  authenticate.mockClear();
  setAsync.mockImplementation((_html, _options, cb) => cb({ status: 'failed' }));
  expect(button('show-original')).not.toBeNull();
  button('show-original').click();
  await settle();
  expect(authenticate).not.toHaveBeenCalled();
  expect(document.getElementById('status')!.textContent).toContain('原文显示失败');
  expect(button('show-original').disabled).toBe(false);
});

test('ignores unknown initialization data', async () => {
  const { setAsync } = await setup('not JSON');
  expect(setAsync).not.toHaveBeenCalled();
});

test('manual restoration invalidates a pending automatic language detection', async () => {
  const { api, setAsync, handlers } = await setup(undefined, 'always');
  api.mockClear(); setAsync.mockClear();
  let finishDetection!: (value: any) => void;
  api.mockImplementationOnce(() => new Promise(resolve => { finishDetection = resolve; }));
  handlers.itemChanged();
  await settle();
  button('show-original').click();
  await settle();
  finishDetection({ language: 'en', score: 1 });
  await settle();
  expect(api).toHaveBeenCalledTimes(1);
  expect(setAsync).toHaveBeenCalledTimes(1);
  expect(setAsync).toHaveBeenCalledWith('<p>Original</p>', { coercionType: 'html' }, expect.any(Function));
});

test('disables both manual actions during translation and reenables them afterward', async () => {
  const { api } = await setup();
  let finishTranslation!: (value: any) => void;
  api.mockImplementationOnce(() => new Promise(resolve => { finishTranslation = resolve; }));
  button('translate-message').click();
  await settle();
  expect(button('show-original').disabled).toBe(true);
  expect(button('translate-message').disabled).toBe(true);
  finishTranslation({ html: '<p>译文</p>' });
  await settle();
  expect(button('show-original').disabled).toBe(false);
  expect(button('translate-message').disabled).toBe(false);
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
  expect(button('show-original').disabled).toBe(true);
  expect(button('translate-message').disabled).toBe(true);
});

test('can restore the new message while a previous message translation is still pending', async () => {
  const { api, handlers, itemHandlers, item, setAsync } = await setup();
  let finishTranslation!: (value: any) => void;
  api.mockImplementationOnce(() => new Promise(resolve => { finishTranslation = resolve; }));
  button('translate-message').click();
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

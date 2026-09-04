jest.mock('../shared/api', () => ({ authenticate: jest.fn(), api: jest.fn() }));
import { authenticate, api } from '../shared/api';
import * as mail from '../shared/mail';
const { translateCurrentMessage } = mail;
function setup() {
  const setAsync = jest.fn((_html, _options, cb) => cb({ status: 'succeeded' }));
  const item = { itemId: 'a', body: { getAsync: jest.fn((_type, cb) => cb({ status: 'succeeded', value: '<p>Hello</p>' })) },
    display: { body: { setAsync } }, notificationMessages: { replaceAsync: jest.fn((_key, _value, cb) => cb?.({ status: 'succeeded' })) } };
  (globalThis as any).Office = { AsyncResultStatus: { Succeeded: 'succeeded' }, CoercionType: { Html: 'html' },
    context: { mailbox: { item } }, MailboxEnums: { ItemNotificationMessageType: { InformationalMessage: 'info', ErrorMessage: 'error', InsightMessage: 'insightMessage' } } };
  (authenticate as jest.Mock).mockResolvedValue({ token: 'token', user: { id: 'u' } });
  (api as jest.Mock).mockResolvedValue({ html: '<p>你好</p>' });
  return { item, setAsync };
}
test('authenticates, translates the entire HTML and replaces only the displayed body', async () => {
  const { setAsync } = setup();
  await translateCurrentMessage('zh-Hans');
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Hello</p>', to: 'zh-Hans' });
  expect(setAsync).toHaveBeenCalledWith('<p>你好</p>', { coercionType: 'html' }, expect.any(Function));
});

test('reports each translation stage so a stalled host call can be located', async () => {
  const { item } = setup();
  await translateCurrentMessage('zh-Hans');
  expect(item.notificationMessages.replaceAsync.mock.calls.map(call => (call[1] as any).message)).toEqual([
    '正在登录…', '正在读取邮件正文…', '正在翻译整封邮件…', '正在显示译文…', '翻译完成。',
  ]);
});
test('keeps the original on translation or SSO failure', async () => {
  const { setAsync } = setup();
  (api as jest.Mock).mockRejectedValue(new Error('network'));
  await expect(translateCurrentMessage('en')).rejects.toThrow('network');
  expect(setAsync).not.toHaveBeenCalled();
  (authenticate as jest.Mock).mockRejectedValue(new Error('SSO'));
  await expect(translateCurrentMessage('en')).rejects.toThrow('SSO');
  expect(setAsync).not.toHaveBeenCalled();
});
test('never replaces another message after the user switches during translation', async () => {
  const { setAsync } = setup();
  (api as jest.Mock).mockImplementation(async () => {
    (Office.context.mailbox as any).item = { itemId: 'b' };
    return { html: '<p>你好</p>' };
  });
  await expect(translateCurrentMessage('en')).rejects.toThrow('邮件已切换');
  expect(setAsync).not.toHaveBeenCalled();
});
test('checks preview capability before authentication or sending the message', async () => {
  const { item } = setup();
  delete (item as any).display;
  await expect(translateCurrentMessage('en')).rejects.toThrow('不支持');
  expect(authenticate).not.toHaveBeenCalled();
});

test('offers a show-original action when translation finishes', async () => {
  const { item } = setup();
  await translateCurrentMessage('zh-Hans');
  expect(item.notificationMessages.replaceAsync).toHaveBeenLastCalledWith(
    'mail-translation-status', expect.objectContaining({
      actions: [{ actionText: '显示原文', actionType: 'showTaskPane', commandId: 'Translation.Options', contextData: JSON.stringify({ action: 'showOriginal' }) }],
    }), expect.any(Function));
});

test('restores the original body without authentication or translation', async () => {
  const { setAsync } = setup();
  await mail.showOriginalMessage();
  expect(setAsync).toHaveBeenCalledWith('<p>Hello</p>', { coercionType: 'html' }, expect.any(Function));
  expect(authenticate).not.toHaveBeenCalled();
  expect(api).not.toHaveBeenCalled();
});

test('points to translation options when Outlook rejects notification actions', async () => {
  const { item } = setup();
  item.notificationMessages.replaceAsync.mockImplementation((_key, value, cb) => {
    cb?.({ status: value.actions ? 'failed' : 'succeeded' });
  });
  await translateCurrentMessage('zh-Hans');
  expect(item.notificationMessages.replaceAsync).toHaveBeenLastCalledWith(
    'mail-translation-status', expect.objectContaining({
      type: 'info', message: '翻译完成。请打开“翻译选项”，点击“显示原文”恢复原文。',
    }), expect.any(Function));
});

test('does not restore into a different message if selection changes during the read', async () => {
  const { item, setAsync } = setup();
  item.body.getAsync.mockImplementation((_type, cb) => {
    (Office.context.mailbox as any).item = { itemId: 'b' };
    cb({ status: 'succeeded', value: '<p>Hello</p>' });
  });
  await expect(mail.showOriginalMessage()).rejects.toThrow('邮件已切换');
  expect(setAsync).not.toHaveBeenCalled();
});

test('reports a display failure when restoring the original', async () => {
  const { setAsync } = setup();
  setAsync.mockImplementation((_html, _options, cb) => cb({ status: 'failed' }));
  await expect(mail.showOriginalMessage()).rejects.toThrow('原文显示失败');
});

test('times out an unresponsive display API, completes the command, and allows retry', async () => {
  jest.useFakeTimers();
  try {
    const { item, setAsync } = setup();
    setAsync.mockImplementation(() => {});
    const { completeCommand } = await import('./command-logic');
    const event = { completed: jest.fn() };
    const result = completeCommand(event, () => translateCurrentMessage('zh-Hans'));
    const assertion = expect(result).rejects.toThrow('由世纪互联运营的Outlook on the Web目前还不支持该接口，请使用Outlook客户端体验该功能。');
    await jest.advanceTimersByTimeAsync(15000);
    await assertion;
    expect(event.completed).toHaveBeenCalledTimes(1);
    expect(item.notificationMessages.replaceAsync).toHaveBeenLastCalledWith('mail-translation-status',
      expect.objectContaining({ type: 'error', message: '由世纪互联运营的Outlook on the Web目前还不支持该接口，请使用Outlook客户端体验该功能。' }), expect.any(Function));
    setAsync.mockImplementation((_html, _options, cb) => cb({ status: 'succeeded' }));
    await expect(translateCurrentMessage('zh-Hans')).resolves.toBeUndefined();
  } finally { jest.useRealTimers(); }
});

test('times out when Outlook never returns the original body', async () => {
  jest.useFakeTimers();
  try {
    const { item, setAsync } = setup();
    item.body.getAsync.mockImplementation(() => {});
    const assertion = expect(translateCurrentMessage('zh-Hans')).rejects.toThrow('读取邮件正文超时');
    await jest.advanceTimersByTimeAsync(15000);
    await assertion;
    expect(api).not.toHaveBeenCalled();
    expect(setAsync).not.toHaveBeenCalled();
  } finally { jest.useRealTimers(); }
});

test('times out restoring the original instead of locking the pane indefinitely', async () => {
  jest.useFakeTimers();
  try {
    const { setAsync } = setup();
    setAsync.mockImplementation(() => {});
    const assertion = expect(mail.showOriginalMessage()).rejects.toThrow('显示原文超时');
    await jest.advanceTimersByTimeAsync(15000);
    await assertion;
  } finally { jest.useRealTimers(); }
});

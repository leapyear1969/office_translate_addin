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

import { api, authenticate, Session } from './api';
import { LANGUAGES, loadSettings } from './settings';
import { getDisplayedBody } from '../commands/command-logic';

export function currentItem(): Office.MessageRead | null {
  return (Office.context.mailbox.item as Office.MessageRead | undefined) || null;
}
export function isCurrent(item: Office.MessageRead): boolean {
  const selected = currentItem();
  return !!selected && (item.itemId ? item.itemId === selected.itemId : item === selected);
}
function officeCall<T>(invoke: (callback: (result: Office.AsyncResult<T>) => void) => void,
  failureMessage: string, timeoutMessage: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(timeoutMessage)), 15000);
    try {
      invoke(result => {
        clearTimeout(timer);
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(`${failureMessage}${result.error?.code ? `（${result.error.code}）` : ''}`));
      });
    } catch (error) { clearTimeout(timer); reject(error); }
  });
}
export function bodyHtml(item: Office.MessageRead): Promise<string> {
  return officeCall(callback => item.body.getAsync(Office.CoercionType.Html, callback),
    '读取邮件正文失败，请重新打开邮件。', '读取邮件正文超时，请重新打开邮件后重试。');
}
export function notify(item: Office.MessageRead, message: string, error = false): void {
  if (!isCurrent(item)) return;
  const notification: Office.NotificationMessageDetails = error
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: message.slice(0, 150) }
    : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: message.slice(0, 150), icon: 'Icon.16', persistent: false };
  item.notificationMessages.replaceAsync('mail-translation-status', notification, () => {});
}
export function notifyConsentRequired(item: Office.MessageRead): void {
  if (!isCurrent(item)) return;
  item.notificationMessages.replaceAsync('mail-translation-status', {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: '首次使用需要授权，请打开翻译选项完成登录。', icon: 'Icon.16',
    actions: [{ actionText: '登录并授权', actionType: 'showTaskPane', commandId: 'Translation.Options',
      contextData: JSON.stringify({ action: 'consent' }) }],
  }, result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) notify(item, '首次使用需要授权，请打开“翻译选项”，点击“登录并授权”。');
  });
}
const running = new Set<string | Office.MessageRead>();

export async function showOriginalMessage(target?: string): Promise<void> {
  const item = currentItem();
  if (!item) throw new Error('请先选择一封邮件。');
  const displayed = getDisplayedBody(item);
  if (!displayed) throw new Error('当前 Outlook 不支持恢复显示，请重新打开邮件查看原文。');
  if (running.has(item.itemId || item)) throw new Error('当前邮件正在翻译，请稍候。');
  // Translation changes display.body only. body.getAsync still reads the original,
  // including when this action starts in a fresh task-pane runtime.
  const html = await bodyHtml(item);
  if (!isCurrent(item)) throw new Error('邮件已切换，请在当前邮件上重新点击“显示原文”。');
  await officeCall<void>(callback => displayed.setAsync(html, { coercionType: Office.CoercionType.Html }, callback),
    '原文显示失败，请重试或重新打开邮件。', '显示原文超时：Outlook 正文显示接口未响应，请重新打开邮件查看原文。');
  notifyOriginalDisplayed(item, target);
}

function notifyMessageAction(item: Office.MessageRead, message: string, actionText: string,
  action: 'showOriginal' | 'translateMessage', fallback: string): void {
  if (!isCurrent(item)) return;
  try {
    item.notificationMessages.replaceAsync('mail-translation-status', {
      type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      message, icon: 'Icon.16',
      actions: [{ actionText, actionType: 'showTaskPane', commandId: 'Translation.Options', contextData: JSON.stringify({ action }) }],
    }, result => {
      // Some Outlook clients cannot show actionable notifications in read mode.
      if (result.status !== Office.AsyncResultStatus.Succeeded) notify(item, fallback);
    });
  } catch { notify(item, fallback); }
}
export function notifyOriginalDisplayed(item: Office.MessageRead, target?: string): void {
  if (!isCurrent(item)) return;
  const actionText = `将邮件翻译为：${LANGUAGES[target || loadSettings().target]}`;
  notifyMessageAction(item, '已显示原文。', actionText, 'translateMessage',
    '已显示原文。请点击功能区中的“翻译邮件”重新翻译。');
}
function notifyTranslationComplete(item: Office.MessageRead): void {
  notifyMessageAction(item, '翻译完成。', '显示原文', 'showOriginal',
    '翻译完成。请重新打开邮件查看原文。');
}
export async function translateCurrentMessage(target: string, session?: Session, expectedItem?: Office.MessageRead): Promise<void> {
  const item = expectedItem || currentItem();
  if (!item) throw new Error('请先选择一封邮件。');
  if (!isCurrent(item)) throw new Error('邮件已切换，请重新翻译。');
  const displayed = getDisplayedBody(item);
  if (!displayed) throw new Error('当前 Outlook 不支持临时替换邮件正文，请使用已验证支持预览 API 的客户端。');
  const key = item.itemId || item;
  if (running.has(key)) throw new Error('当前邮件正在翻译，请稍候。');
  running.add(key);
  try {
    notify(item, '正在登录…');
    const authenticated = session || await authenticate();
    if (!isCurrent(item)) throw new Error('邮件已切换，已取消应用译文。');
    notify(item, '正在读取邮件正文…');
    const html = await bodyHtml(item);
    if (!html.trim()) throw new Error('这封邮件没有可翻译的正文。');
    notify(item, '正在翻译整封邮件…');
    const result = await api<{ html: string }>('/api/translate', authenticated.token, { html, to: target });
    if (!isCurrent(item)) throw new Error('邮件已切换，已取消应用译文。');
    if (typeof result.html !== 'string' || !result.html || result.html.length > 1000000) throw new Error('译文无效或超过显示限制。');
    notify(item, '正在显示译文…');
    await officeCall<void>(callback => displayed.setAsync(result.html, { coercionType: Office.CoercionType.Html }, callback),
      '译文显示失败，原始邮件未修改。',
      '由世纪互联运营的Outlook on the Web目前还不支持该接口，请使用Outlook客户端体验该功能。');
    notifyTranslationComplete(item);
  } catch (error) {
    if ((error as { code?: string }).code === 'consent_required') notifyConsentRequired(item);
    else notify(item, error instanceof Error ? error.message : '翻译失败，请重试。', true);
    throw error;
  } finally { running.delete(key); }
}

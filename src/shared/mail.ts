import { api, authenticate, Session } from './api';
import { getDisplayedBody } from '../commands/command-logic';

export function currentItem(): Office.MessageRead | null {
  return (Office.context.mailbox.item as Office.MessageRead | undefined) || null;
}
export function isCurrent(item: Office.MessageRead): boolean {
  const selected = currentItem();
  return !!selected && (item.itemId ? item.itemId === selected.itemId : item === selected);
}
export function bodyHtml(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve, reject) => item.body.getAsync(Office.CoercionType.Html, result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
    else reject(new Error('读取邮件正文失败，请重新打开邮件。'));
  }));
}
export function notify(item: Office.MessageRead, message: string, error = false): void {
  if (!isCurrent(item)) return;
  const notification: Office.NotificationMessageDetails = error
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: message.slice(0, 150) }
    : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: message.slice(0, 150), icon: 'Icon.16', persistent: false };
  item.notificationMessages.replaceAsync('mail-translation-status', notification, () => {});
}
const running = new Set<string | Office.MessageRead>();

export async function showOriginalMessage(): Promise<void> {
  const item = currentItem();
  if (!item) throw new Error('请先选择一封邮件。');
  const displayed = getDisplayedBody(item);
  if (!displayed) throw new Error('当前 Outlook 不支持恢复显示，请重新打开邮件查看原文。');
  if (running.has(item.itemId || item)) throw new Error('当前邮件正在翻译，请稍候。');
  // Translation changes display.body only. body.getAsync still reads the original,
  // including when this action starts in a fresh task-pane runtime.
  const html = await bodyHtml(item);
  if (!isCurrent(item)) throw new Error('邮件已切换，请在当前邮件上重新点击“显示原文”。');
  await new Promise<void>((resolve, reject) => displayed.setAsync(html, { coercionType: Office.CoercionType.Html }, result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
    else reject(new Error('原文显示失败，请重试或重新打开邮件。'));
  }));
  notify(item, '已显示原文。');
}

function notifyTranslationComplete(item: Office.MessageRead): void {
  if (!isCurrent(item)) return;
  item.notificationMessages.replaceAsync('mail-translation-status', {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: '翻译完成。',
    icon: 'Icon.16',
    actions: [{ actionText: '显示原文', actionType: 'showTaskPane', commandId: 'Translation.Options', contextData: JSON.stringify({ action: 'showOriginal' }) }],
  }, result => {
    // Some Outlook clients cannot show actionable notifications in read mode.
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      notify(item, '翻译完成。请打开“翻译选项”，点击“显示原文”恢复原文。');
    }
  });
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
    notify(item, '正在登录并翻译整封邮件…');
    const authenticated = session || await authenticate();
    if (!isCurrent(item)) throw new Error('邮件已切换，已取消应用译文。');
    const html = await bodyHtml(item);
    if (!html.trim()) throw new Error('这封邮件没有可翻译的正文。');
    const result = await api<{ html: string }>('/api/translate', authenticated.token, { html, to: target });
    if (!isCurrent(item)) throw new Error('邮件已切换，已取消应用译文。');
    if (typeof result.html !== 'string' || !result.html || result.html.length > 1000000) throw new Error('译文无效或超过显示限制。');
    await new Promise<void>((resolve, reject) => displayed.setAsync(result.html, { coercionType: Office.CoercionType.Html }, response => {
      if (response.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error('译文显示失败，原始邮件未修改。'));
    }));
    notifyTranslationComplete(item);
  } finally { running.delete(key); }
}

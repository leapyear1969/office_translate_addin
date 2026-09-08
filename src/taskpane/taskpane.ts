import { setupAccount } from '../shared/account';
import { api, authenticate, Session } from '../shared/api';
import { requestConsent } from '../shared/consent';
import { bodyHtml, currentItem, isCurrent, notifyOriginalDisplayed, showOriginalMessage, translateCurrentMessage } from '../shared/mail';
import { LANGUAGES, loadSettings, saveSettings, Settings, shouldOfferTranslation } from '../shared/settings';

const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const target = element<HTMLSelectElement>('target-language');
const excluded = element<HTMLSelectElement>('excluded-languages');
const picker = element<HTMLSelectElement>('new-language');
let settings: Settings;
let session: Session | undefined;
let epoch = 0;
let loginEpoch = 0;
let actionEpoch = 0;
let offeredItem: Office.MessageRead | undefined;
let translating = false;
let restoring = false;
let preserveOriginal = false;
let dirty = false;
let authorizing = false;
let signedOut = false;
const account = setupAccount(() => {
  signedOut = true;
  ++loginEpoch;
  ++epoch;
  session = undefined;
  offeredItem = undefined;
  element('translation-prompt').hidden = true;
  element('browser-consent').hidden = true;
  element('account-name').textContent = '已注销';
  element('account-email').textContent = '';
  element('signin').hidden = false;
  element('signin').textContent = '重新登录';
  element<HTMLButtonElement>('signin').disabled = false;
  account.render();
  updateMessageActions();
  status('已注销当前面板账户。点击头像可重新登录。');
});
function updateMessageActions() {
  element('translate-message').textContent = `将邮件翻译为：${LANGUAGES[settings.target]}`;
  const disabled = translating || restoring || !currentItem();
  ['show-original', 'translate-message', 'translate-now'].forEach(id => { element<HTMLButtonElement>(id).disabled = disabled || (signedOut && id !== 'show-original'); });
}
function refreshOriginalNotification() {
  const item = currentItem();
  if (!item || translating || restoring) return;
  const ownAction = actionEpoch;
  const savedTarget = settings.target;
  try {
    item.notificationMessages.getAllAsync(result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !isCurrent(item)
        || ownAction !== actionEpoch || savedTarget !== settings.target) return;
      // A ribbon command can change the display independently of this pane.
      const notification = result.value.find(value => value.key === 'mail-translation-status');
      if (notification?.message.startsWith('已显示原文。')) notifyOriginalDisplayed(item, savedTarget);
    });
  } catch { /* A label refresh must not prevent settings from being saved. */ }
}
function status(message: string, error = false) {
  element('status').textContent = message;
  element('status').classList.toggle('error', error);
}
function fillLanguages(select: HTMLSelectElement) {
  Object.entries(LANGUAGES).forEach(([code, label]) => select.add(new Option(label, code)));
}
function renderExcluded(codes: string[]) {
  excluded.replaceChildren(...codes.map(code => new Option(LANGUAGES[code], code)));
  element('empty-languages').hidden = codes.length > 0;
  element<HTMLButtonElement>('remove-language').disabled = excluded.selectedIndex < 0;
}
function markDirty() { dirty = true; element('save-status').textContent = '尚未保存'; }
function currentForm(): Settings {
  return { mode: (document.querySelector<HTMLInputElement>('input[name=mode]:checked')!.value as Settings['mode']),
    target: target.value, excluded: Array.from(excluded.options).map(option => option.value) };
}
async function inspectMessage() {
  const ownEpoch = ++epoch;
  offeredItem = undefined;
  element('translation-prompt').hidden = true;
  const item = currentItem();
  if (!session || !item || settings.mode === 'never' || preserveOriginal) return;
  try {
    status('正在识别邮件语言…');
    const html = await bodyHtml(item);
    if (ownEpoch !== epoch || !isCurrent(item)) return;
    const detected = await api<{ language: string; score: number } | null>('/api/detect', session.token, { html });
    if (ownEpoch !== epoch || !isCurrent(item)) return;
    if (!detected || !shouldOfferTranslation(settings, detected.language, detected.score)) { status('当前邮件无需自动翻译。'); return; }
    offeredItem = item;
    if (settings.mode === 'always') await translateOffered();
    else {
      element('prompt-message').textContent = `这封邮件使用${LANGUAGES[detected.language] || detected.language}，是否翻译为${LANGUAGES[settings.target]}？`;
      element('translation-prompt').hidden = false;
      status('');
    }
  } catch (error) { if (ownEpoch === epoch) status((error as Error).message, true); }
}
async function translateOffered() {
  if (offeredItem) await translateItem(offeredItem);
}
async function translateItem(item: Office.MessageRead | null) {
  if (!item || translating || restoring || signedOut) return;
  const ownEpoch = ++epoch;
  const ownAction = ++actionEpoch;
  preserveOriginal = false;
  translating = true;
  updateMessageActions();
  try {
    status('正在翻译整封邮件…');
    // Reacquire SSO for each translation rather than retain expired access tokens.
    await translateCurrentMessage(settings.target, undefined, item);
    if (ownEpoch === epoch) { element('translation-prompt').hidden = true; status('翻译完成。点击提示栏或此面板中的“显示原文”即可恢复原文。'); }
  } catch (error) { if (ownEpoch === epoch) status((error as Error).message, true); }
  finally { if (ownAction === actionEpoch) { translating = false; updateMessageActions(); } }
}
async function restoreOriginal() {
  if (translating || restoring) return;
  const item = currentItem();
  const restoreTarget = settings.target;
  const ownEpoch = ++epoch;
  const ownAction = ++actionEpoch;
  preserveOriginal = true;
  offeredItem = undefined;
  element('translation-prompt').hidden = true;
  restoring = true;
  updateMessageActions();
  status('正在显示原文…');
  try {
    await showOriginalMessage(restoreTarget);
    if (ownAction === actionEpoch && item && isCurrent(item) && restoreTarget !== settings.target) {
      notifyOriginalDisplayed(item, settings.target);
    }
    if (ownEpoch === epoch) status('已显示原文。');
  } catch (error) { if (ownEpoch === epoch) status((error as Error).message, true); }
  finally { if (ownAction === actionEpoch) { restoring = false; updateMessageActions(); } }
}
async function handleInitializationContext(data: unknown): Promise<boolean> {
  try {
    const context = typeof data === 'string' ? JSON.parse(data) : data;
    if (context?.action === 'showOriginal') { await restoreOriginal(); return true; }
    if (context?.action === 'translateMessage') { await translateItem(currentItem()); return true; }
  } catch { /* Empty or unrelated launch data is not a mail action. */ }
  return false;
}
function registerInitializationHandler() {
  const item = currentItem();
  if (!item || !Office.EventType.InitializationContextChanged || typeof item.addHandlerAsync !== 'function') return;
  item.addHandlerAsync(Office.EventType.InitializationContextChanged, (event: Office.InitializationContextChangedEventArgs) => {
    if (isCurrent(item)) void handleInitializationContext(event.initializationContextData);
  }, () => {});
}
function readInitializationContext(): Promise<boolean> {
  const item = currentItem();
  const ownAction = actionEpoch;
  if (!item || typeof item.getInitializationContextAsync !== 'function') return Promise.resolve(false);
  return new Promise(resolve => {
    item.getInitializationContextAsync(result => {
      // A newer notification event supersedes delayed launch data, including empty data.
      if (ownAction !== actionEpoch) { resolve(true); return; }
      if (result.status === Office.AsyncResultStatus.Succeeded && isCurrent(item)) {
        void handleInitializationContext(result.value).then(resolve);
      } else resolve(false);
    });
  });
}
async function signIn(interactive: boolean, inspect = true) {
  if (authorizing || (signedOut && !interactive)) return;
  account.close();
  const attempt = ++loginEpoch;
  const ownAction = actionEpoch;
  element<HTMLButtonElement>('signin').disabled = true;
  element('account-name').textContent = '正在读取登录账户…';
  element('browser-consent').hidden = true;
  try {
    let authenticated: Session;
    try { authenticated = await authenticate(interactive); }
    catch (error) {
      const failure = error as { code?: string; token?: string };
      if (!interactive || failure.code !== 'consent_required' || !failure.token) throw error;
      authorizing = true;
      const url = await requestConsent(failure.token);
      if (attempt !== loginEpoch) return;
      element<HTMLAnchorElement>('browser-consent-link').href = url;
      element('browser-consent').hidden = false;
      session = undefined;
      account.render();
      element('account-name').textContent = '等待浏览器授权';
      element('account-email').textContent = '';
      element('signin').hidden = false;
      element('signin').textContent = '登录并授权';
      status('请在浏览器中使用当前 Outlook 账户完成授权，然后关闭此窗口并重新打开插件。');
      return;
    }
    if (attempt !== loginEpoch) return;
    session = authenticated;
    signedOut = false;
    account.render(session.user);
    updateMessageActions();
    element('account-name').textContent = session.user.displayName || '已登录';
    element('account-email').textContent = session.user.mail;
    element('signin').hidden = false;
    element('signin').textContent = '重新登录';
    if (inspect && !preserveOriginal && ownAction === actionEpoch) { status(''); await inspectMessage(); }
  } catch (error) {
    if (attempt !== loginEpoch) return;
    session = undefined;
    account.render();
    element('account-name').textContent = '尚未完成登录';
    element('account-email').textContent = '';
    element('signin').hidden = false;
    element('signin').textContent = '登录并授权';
    if (inspect && !preserveOriginal && ownAction === actionEpoch) status((error as Error).message, true);
  } finally { authorizing = false; if (attempt === loginEpoch) element<HTMLButtonElement>('signin').disabled = false; }
}

fillLanguages(target); fillLanguages(picker); target.value = 'zh-Hans';
element('signin').addEventListener('click', () => void signIn(true));
element('preferences').addEventListener('change', event => {
  if (event.target !== excluded && event.target !== picker) markDirty();
});
excluded.addEventListener('change', () => { element<HTMLButtonElement>('remove-language').disabled = excluded.selectedIndex < 0; });
element('add-language').addEventListener('click', () => {
  element('language-picker').hidden = false;
  Array.from(picker.options).forEach(option => { option.disabled = Array.from(excluded.options).some(existing => existing.value === option.value); });
  picker.value = Array.from(picker.options).find(option => !option.disabled)?.value || '';
  element<HTMLButtonElement>('confirm-language').disabled = !picker.value;
  picker.focus();
});
element('cancel-language').addEventListener('click', () => { element('language-picker').hidden = true; element('add-language').focus(); });
element('confirm-language').addEventListener('click', () => {
  const codes = currentForm().excluded;
  if (picker.value && !codes.includes(picker.value)) { renderExcluded([...codes, picker.value]); markDirty(); }
  element('language-picker').hidden = true;
  element('add-language').focus();
});
element('remove-language').addEventListener('click', () => {
  if (excluded.selectedIndex >= 0) { renderExcluded(currentForm().excluded.filter(code => code !== excluded.value)); markDirty(); }
});
element('preferences').addEventListener('submit', async event => {
  event.preventDefault();
  if (!dirty) { element('save-status').textContent = '设置已保存'; return; }
  const fields = element<HTMLFieldSetElement>('settings-fields');
  fields.disabled = true;
  try {
    const next = currentForm();
    await saveSettings(next); settings = next; dirty = false;
    updateMessageActions();
    refreshOriginalNotification();
    element('save-status').textContent = '设置已保存';
    await inspectMessage();
  } catch (error) { element('save-status').textContent = (error as Error).message; }
  finally { fields.disabled = false; }
});
element('translate-now').addEventListener('click', () => void translateOffered());
element('translate-message').addEventListener('click', () => void translateItem(currentItem()));
element('show-original').addEventListener('click', () => void restoreOriginal());
element('dismiss').addEventListener('click', () => { offeredItem = undefined; element('translation-prompt').hidden = true; });

Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) return;
  settings = loadSettings();
  target.value = settings.target;
  document.querySelector<HTMLInputElement>(`input[name=mode][value=${settings.mode}]`)!.checked = true;
  renderExcluded(settings.excluded);
  updateMessageActions();
  element<HTMLFieldSetElement>('settings-fields').disabled = false;
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    ++epoch;
    // An old item's in-flight operation must not lock or unlock the new item's UI.
    ++actionEpoch;
    translating = false;
    restoring = false;
    preserveOriginal = false;
    offeredItem = undefined;
    element('translation-prompt').hidden = true;
    updateMessageActions();
    registerInitializationHandler();
    void signIn(false);
  }, result => { if (result.status !== Office.AsyncResultStatus.Succeeded) status('当前客户端无法监听邮件切换，请重新打开面板以识别新邮件。'); });
  registerInitializationHandler();
  const initialLogin = loginEpoch;
  void readInitializationContext().then(handled => {
    if (initialLogin === loginEpoch) return signIn(false, !handled);
  });
});

import { setupAccount } from '../shared/account';
import { api, authenticate, Session } from '../shared/api';
import { requestConsent } from '../shared/consent';
import { documentHtml, translateDocument, restoreOriginalBody, Scope } from './document';
import { loadSettings, saveSettings } from './settings';
import { LANGUAGES, Settings, shouldOfferTranslation } from '../shared/settings';

const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const target = element<HTMLSelectElement>('target-language');
const excluded = element<HTMLSelectElement>('excluded-languages');
const picker = element<HTMLSelectElement>('new-language');
let settings: Settings;
let session: Session | undefined;
let epoch = 0;
let loginEpoch = 0;
let offeredHtml: string | undefined;
let translating = false;
let dirty = false;
let authorizing = false;
let visible = false;
let initialized = false;
let suppressAutomatic = false;
let signedOut = false;
const account = setupAccount(() => {
  signedOut = true;
  ++loginEpoch;
  ++epoch;
  session = undefined;
  offeredHtml = undefined;
  element('translation-prompt').hidden = true;
  element('browser-consent').hidden = true;
  element('account-name').textContent = '已注销';
  element('account-email').textContent = '';
  element('signin').hidden = false;
  element('signin').textContent = '重新登录';
  element<HTMLButtonElement>('signin').disabled = false;
  account.render();
  updateDocumentActions();
  status('已注销当前面板账户。点击头像可重新登录。');
});
function updateDocumentActions() {
  ['translate-selection', 'translate-paragraph', 'translate-body', 'translate-now', 'restore-original', 'confirm-restore'].forEach(id => {
    element<HTMLButtonElement>(id).disabled = translating || (signedOut && id.startsWith('translate'));
  });
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
async function inspectDocument() {
  const ownEpoch = ++epoch;
  offeredHtml = undefined;
  element('translation-prompt').hidden = true;
  if (!visible || !session || translating || suppressAutomatic || settings.mode === 'never') return;
  try {
    status('正在识别文档语言…');
    const html = await documentHtml();
    const detected = await api<{ language: string; score: number } | null>('/api/detect', session.token, { html });
    if (ownEpoch !== epoch || !visible || translating) return;
    if (!detected || !shouldOfferTranslation(settings, detected.language, detected.score)) { status('当前文档无需自动翻译。'); return; }
    offeredHtml = html;
    if (settings.mode === 'always') await translateScope('body', settings.target, html);
    else {
      element('prompt-message').textContent = `此文档使用${LANGUAGES[detected.language] || detected.language}，是否翻译为${LANGUAGES[settings.target]}？`;
      element('translation-prompt').hidden = false;
      status('');
    }
  } catch (error) { if (ownEpoch === epoch) status((error as Error).message, true); }
}
async function translateScope(scope: Scope, targetLanguage = settings.target, expectedHtml?: string) {
  if (translating || authorizing || signedOut) return;
  if (expectedHtml === undefined) suppressAutomatic = true;
  const actionEpoch = ++epoch;
  translating = true;
  updateDocumentActions();
  element('translation-prompt').hidden = true;
  element('restore-prompt').hidden = true;
  status('正在翻译文档内容…');
  try {
    await translateDocument(scope, targetLanguage, expectedHtml, () => !signedOut && (expectedHtml === undefined || (visible && epoch === actionEpoch)));
    status('翻译完成，原文备份已保留。保存文档后，重新打开也可恢复原文。');
  } catch (error) {
    status((error as Error).message, true);
    if ((error as { code?: string }).code === 'consent_required') {
      element('signin').hidden = false;
      element('signin').textContent = '登录并授权';
    }
  } finally { translating = false; updateDocumentActions(); }
}
async function signIn(interactive: boolean) {
  if (authorizing || (signedOut && !interactive)) return;
  account.close();
  const attempt = ++loginEpoch;
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
      status('请在浏览器中使用当前 Word 账户完成授权，然后关闭此窗口并重新打开插件。');
      return;
    }
    if (attempt !== loginEpoch) return;
    session = authenticated;
    signedOut = false;
    account.render(session.user);
    updateDocumentActions();
    element('account-name').textContent = session.user.displayName || '已登录';
    element('account-email').textContent = session.user.mail;
    element('signin').hidden = false;
    element('signin').textContent = '重新登录';
    if (!translating && !suppressAutomatic) { status(''); await inspectDocument(); }
  } catch (error) {
    if (attempt !== loginEpoch) return;
    session = undefined;
    account.render();
    element('account-name').textContent = '尚未完成登录';
    element('account-email').textContent = '';
    element('signin').hidden = false;
    element('signin').textContent = '登录并授权';
    if (!translating) status((error as Error).message, true);
  } finally { authorizing = false; if (attempt === loginEpoch) element<HTMLButtonElement>('signin').disabled = false; }
}

fillLanguages(target); fillLanguages(picker); target.value = 'zh-Hans';
element('restore-original').addEventListener('click', () => {
  ++epoch;
  suppressAutomatic = true;
  offeredHtml = undefined;
  element('translation-prompt').hidden = true;
  element('restore-prompt').hidden = false;
});
element('cancel-restore').addEventListener('click', () => { element('restore-prompt').hidden = true; });
element('confirm-restore').addEventListener('click', async () => {
  if (translating || authorizing) return;
  ++epoch;
  suppressAutomatic = true;
  translating = true;
  updateDocumentActions();
  element('restore-prompt').hidden = true;
  status('正在恢复原文…');
  try {
    await restoreOriginalBody();
    status('已恢复首次翻译前的正文，原文备份继续保留。请保存文档。');
  } catch (error) { status((error as Error).message, true); }
  finally { translating = false; updateDocumentActions(); }
});
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
    await saveSettings(next); settings = next; dirty = false; suppressAutomatic = false;
    updateDocumentActions();
    element('save-status').textContent = '设置已保存';
    await inspectDocument();
  } catch (error) { element('save-status').textContent = (error as Error).message; }
  finally { fields.disabled = false; }
});
element('translate-now').addEventListener('click', () => {
  if (offeredHtml !== undefined) void translateScope('body', settings.target, offeredHtml);
});
(['selection', 'paragraph', 'body'] as Scope[]).forEach(scope => {
  element(`translate-${scope}`).addEventListener('click', () => void translateScope(scope));
});
element('dismiss').addEventListener('click', () => { ++epoch; offeredHtml = undefined; element('translation-prompt').hidden = true; });

Office.actions.associate('translateSelectionChinese', async (event: Office.AddinCommands.Event) => {
  try {
    // Start capture immediately: showing the pane or SSO may move focus.
    const translation = translateScope('selection', 'zh-Hans');
    await Office.addin.showAsTaskpane();
    await translation;
  } catch (error) { status((error as Error).message, true); }
  finally { event.completed(); }
});

Office.onReady(async info => {
  if (info.host !== Office.HostType.Word) return;
  settings = loadSettings();
  target.value = settings.target;
  document.querySelector<HTMLInputElement>(`input[name=mode][value=${settings.mode}]`)!.checked = true;
  renderExcluded(settings.excluded);
  updateDocumentActions();
  element<HTMLFieldSetElement>('settings-fields').disabled = false;
  await Office.addin.onVisibilityModeChanged(event => {
    visible = event.visibilityMode === Office.VisibilityMode.taskpane;
    if (!visible) { ++epoch; element('translation-prompt').hidden = true; }
    else if (initialized && !translating) {
      if (session) void inspectDocument();
      else void signIn(false);
    }
  });
  await Office.addin.showAsTaskpane();
  visible = true;
  initialized = true;
  void signIn(false);
});

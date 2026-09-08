import { setupAccount } from '../shared/account';
import { authenticate, Session } from '../shared/api';
import { requestConsent } from '../shared/consent';
import { translateDocument, restoreOriginalBody, Scope } from './document';
import { loadSettings, saveSettings, WordSettings } from './settings';
import { LANGUAGES } from '../shared/settings';
import { isWordOnline } from './web-safety';

const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const target = element<HTMLSelectElement>('target-language');
let settings: WordSettings;
let session: Session | undefined;
let loginEpoch = 0;
let translating = false;
let dirty = false;
let authorizing = false;
let visible = false;
let initialized = false;
let signedOut = false;
const account = setupAccount(() => {
  signedOut = true;
  ++loginEpoch;
  session = undefined;
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
  ['translate-selection', 'translate-paragraph', 'translate-body', 'restore-original', 'confirm-restore'].forEach(id => {
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
function markDirty() { dirty = true; element('save-status').textContent = '尚未保存'; }
function currentForm(): WordSettings { return { target: target.value }; }
async function translateScope(scope: Scope, targetLanguage = settings.target) {
  if (translating || authorizing || signedOut) return;
  translating = true;
  updateDocumentActions();
  element('restore-prompt').hidden = true;
  status(isWordOnline() ? '正在检查文档是否适合网页版翻译…' : '正在翻译文档内容…');
  try {
    const result = await translateDocument(scope, targetLanguage, () => !signedOut);
    status((result?.htmlBackup
      ? '翻译完成。Word 无法导出 OOXML，已使用 HTML 原文备份；可恢复文字及基本格式，复杂格式可能变化。请保存文档。'
      : '翻译完成，原文备份已保留。保存文档后，重新打开也可恢复原文。')
      + (result?.skippedParagraphs ? `有 ${result.skippedParagraphs} 个段落因包含复杂结构或译文标记不匹配而跳过，已保留原文。` : ''));
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
    if (!translating) status('');
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

fillLanguages(target); target.value = 'zh-Hans';
element('restore-original').addEventListener('click', () => {
  element('restore-prompt').hidden = false;
});
element('cancel-restore').addEventListener('click', () => { element('restore-prompt').hidden = true; });
element('confirm-restore').addEventListener('click', async () => {
  if (translating || authorizing) return;
  translating = true;
  updateDocumentActions();
  element('restore-prompt').hidden = true;
  status('正在恢复原文…');
  try {
    const result = await restoreOriginalBody();
    status((result.skipped && result.details?.length
      ? `已恢复 ${result.restored} 处翻译；${result.skipped} 处跳过：${[...new Set(result.details)].join('；')}。已保留当前编辑。`
      : result.skipped
      ? `已恢复 ${result.restored} 处翻译；${result.skipped} 处因文字已修改、标记重复或备份未完成而跳过，已保留当前编辑。`
      : result.restored ? `已恢复 ${result.restored} 处翻译的原文，其他编辑已保留。请保存文档。` : '没有可恢复的翻译范围，当前正文未改变。')
      + (result.cleaned ? `已清理 ${result.cleaned} 个未完成的旧标记，仅移除外框，文字和原文备份保留。请保存文档。` : ''));
  } catch (error) { status((error as Error).message, true); }
  finally { translating = false; updateDocumentActions(); }
});
element('signin').addEventListener('click', () => void signIn(true));
target.addEventListener('change', markDirty);
element('preferences').addEventListener('submit', async event => {
  event.preventDefault();
  if (!dirty) { element('save-status').textContent = '设置已保存'; return; }
  const fields = element<HTMLFieldSetElement>('settings-fields');
  fields.disabled = true;
  try {
    const next = currentForm();
    await saveSettings(next); settings = next; dirty = false;
    updateDocumentActions();
    element('save-status').textContent = '设置已保存';
  } catch (error) { element('save-status').textContent = (error as Error).message; }
  finally { fields.disabled = false; }
});
(['selection', 'paragraph', 'body'] as Scope[]).forEach(scope => {
  element(`translate-${scope}`).addEventListener('click', () => void translateScope(scope));
});

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
  element('web-backup-notice').hidden = !isWordOnline();
  settings = loadSettings();
  target.value = settings.target;
  updateDocumentActions();
  element<HTMLFieldSetElement>('settings-fields').disabled = false;
  await Office.addin.onVisibilityModeChanged(event => {
    visible = event.visibilityMode === Office.VisibilityMode.taskpane;
    if (visible && initialized && !translating && !session) void signIn(false);
  });
  await Office.addin.showAsTaskpane();
  visible = true;
  initialized = true;
  void signIn(false);
});

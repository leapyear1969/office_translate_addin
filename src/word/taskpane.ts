import { setupAccount } from '../shared/account';
import { authenticate, Session } from '../shared/api';
import { requestConsent } from '../shared/consent';
import { translateDocument, restoreOriginalBody, clearTranslationControls, Scope } from './document';
import { loadSettings, saveSettings, WordSettings } from './settings';
import { LANGUAGES } from '../shared/settings';
import { isWordOnline } from './web-safety';
import { migrateDocumentBackups } from './backup';
import { isDesktopWord } from './desktop-copy';
import { capturePreview, translatePreview, PreviewRange, EmptySelectionError } from './preview';

const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const target = element<HTMLSelectElement>('target-language');
let settings: WordSettings = { target: 'zh-Hans' };
let activeScope: 'selection' | 'paragraph' = 'selection';
let session: Session | undefined;
let loginEpoch = 0;
let translating = false;
let authorizing = false;
let visible = false;
let initialized = false;
let signedOut = false;
let previewRange: PreviewRange | undefined;
let previewEpoch = 0;
let previewTimer: ReturnType<typeof setTimeout> | undefined;
let composing = false;
let previewReady = false;
let previewRunning = false;
let queuedPreview: (() => Promise<void>) | undefined;
const sourceText = element<HTMLTextAreaElement>('source-text');
const translatedText = element<HTMLTextAreaElement>('translated-text');
const previewTarget = element<HTMLSelectElement>('preview-target');
// Keep panel sizing local to this browser, independent of document settings.
for (const field of [sourceText, translatedText]) {
  const key = `word-translator.height.${field.id}`;
  try {
    const height = Number(localStorage.getItem(key));
    if (Number.isFinite(height) && height > 0) field.style.height = `${height}px`;
  } catch { /* Storage may be unavailable in the Office webview. */ }
  if (typeof ResizeObserver !== 'undefined') {
    new ResizeObserver(() => {
      // Native textarea resizing sets an inline height. Ignore default layout
      // and hidden panels so switching translation scope cannot erase it.
      const height = parseFloat(field.style.height);
      if (!Number.isFinite(height) || height <= 0) return;
      try { localStorage.setItem(key, String(height)); } catch { /* Keep resizing usable. */ }
    }).observe(field);
  }
}
function invalidatePreview() {
  ++previewEpoch;
  queuedPreview = undefined;
  clearTimeout(previewTimer);
  previewReady = false;
  translatedText.value = '';
  element<HTMLButtonElement>('insert-translation').disabled = true;
}
async function releasePreview() {
  const previous = previewRange;
  previewRange = undefined;
  if (previous) await previous.release();
}
function schedulePreview(immediate = false) {
  invalidatePreview();
  element('source-count').textContent = `（${Array.from(sourceText.value).length} 个字符）`;
  if (signedOut || translating || composing || !sourceText.value.trim()) {
    status(sourceText.value.trim() ? '' : '请输入需要翻译的原文。');
    return;
  }
  const epoch = previewEpoch;
  const text = sourceText.value;
  const language = previewTarget.value;
  status('正在翻译…');
  const run = async () => {
    if (epoch !== previewEpoch || signedOut || translating || composing) return;
    if (previewRunning) { queuedPreview = run; return; }
    previewRunning = true;
    try {
      const result = await translatePreview(text, language);
      if (epoch !== previewEpoch || signedOut) return;
      translatedText.value = result;
      previewReady = true;
      element<HTMLButtonElement>('insert-translation').disabled = !previewRange;
      status('翻译完成');
    } catch (error) {
      if (epoch !== previewEpoch) return;
      status((error as Error).message, true);
    } finally {
      previewRunning = false;
      const next = queuedPreview;
      queuedPreview = undefined;
      if (next) void next();
    }
  };
  if (immediate) void run();
  else previewTimer = setTimeout(() => void run(), 600);
}
async function openPreview(scope: 'selection' | 'paragraph', language: string) {
  activeScope = scope;
  showView(false);
  invalidatePreview();
  const epoch = previewEpoch;
  // Capture before showing the pane, which can change focus in Word.
  const capture = capturePreview(scope);
  try {
    const captured = await capture;
    if (epoch !== previewEpoch || signedOut) { await captured.release(); return; }
    await releasePreview();
    if (epoch !== previewEpoch || signedOut) { await captured.release(); return; }
    previewRange = captured;
    sourceText.value = captured.text;
    previewTarget.value = language;
    showView(false);
    status('');
    schedulePreview(true);
  } catch (error) {
    if (epoch !== previewEpoch) return;
    if (error instanceof EmptySelectionError) {
      sourceText.value = '';
      element('source-count').textContent = '（0 个字符）';
      status(error.message);
      await releasePreview().catch(() => {});
    } else {
      status((error as Error).message, true);
    }
  }
}
function showView(fullDocument: boolean) {
  element('selection-view').hidden = fullDocument;
  element('document-view').hidden = !fullDocument;
  element('translate-selection').setAttribute('aria-selected', String(!fullDocument && activeScope === 'selection'));
  element('translate-paragraph').setAttribute('aria-selected', String(!fullDocument && activeScope === 'paragraph'));
  element('selection-view').setAttribute('aria-labelledby', `translate-${activeScope}`);
  element('document-tab').setAttribute('aria-selected', String(fullDocument));
}
const account = setupAccount(() => {
  invalidatePreview();
  void releasePreview().catch(() => {});
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
  sourceText.disabled = translating || signedOut;
  previewTarget.disabled = translating || signedOut;
  element<HTMLButtonElement>('insert-translation').disabled = translating || signedOut || !previewReady || !previewRange;
  ['translate-selection', 'translate-paragraph', 'translate-body', 'restore-original', 'confirm-restore', 'compact-backups', 'clear-translation-controls'].forEach(id => {
    element<HTMLButtonElement>(id).disabled = translating || (signedOut && id.startsWith('translate'));
  });
}
function status(message: string, error = false, repairFailed = false) {
  const repair = repairFailed || error && /尚未恢复的翻译|翻译控件.*异常|翻译控件.*损坏/.test(message);
  const kind = repair ? 'warning' : error ? 'error' : /正在|请|尚未|没有|已注销/.test(message) ? 'info' : 'success';
  element('notification').hidden = !message;
  element('notification').className = `notification ${kind}`;
  element('status').textContent = repairFailed ? `修复失败：${message}` : repair ? '检测到翻译控件异常，是否修复？' : message;
  element('status').classList.toggle('error', error);
  element('status-icon').className = `bi bi-${kind === 'success' ? 'check-circle-fill' : kind === 'info' ? 'info-circle' : 'exclamation-triangle-fill'}`;
  element('clear-translation-controls').hidden = !repair;
  element('repair-description').hidden = !repair;
}
function fillLanguages(select: HTMLSelectElement) {
  Object.entries(LANGUAGES).forEach(([code, label]) => select.add(new Option(label, code)));
}
async function translateScope(scope: Scope, targetLanguage = settings.target) {
  if (translating || authorizing || signedOut) return;
  if (scope !== 'body') { await openPreview(scope, targetLanguage); return; }
  showView(true);
  invalidatePreview();
  translating = true;
  updateDocumentActions();
  element('restore-prompt').hidden = true;
  status(isWordOnline() ? '正在检查文档是否适合网页版翻译…' : '正在准备翻译副本并翻译…');
  try {
    await releasePreview();
    const result = await translateDocument(scope, targetLanguage, () => !signedOut);
    status((result?.desktopCopy
      ? '翻译已在副本中完成。请保存副本；后续翻译和恢复请在副本中打开插件操作，原文备份随副本保存。'
      : result?.htmlBackup
      ? '翻译完成。Word 无法导出 OOXML，已使用 HTML 原文备份；可恢复文字及基本格式，复杂格式可能变化。请保存文档。'
      : '翻译完成，原文备份已保存到当前浏览器。请保存文档；恢复时需使用同一浏览器和插件地址。')
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
fillLanguages(previewTarget); previewTarget.value = 'zh-Hans';
sourceText.addEventListener('input', () => schedulePreview());
sourceText.addEventListener('compositionstart', () => { composing = true; invalidatePreview(); });
sourceText.addEventListener('compositionend', () => { composing = false; schedulePreview(); });
previewTarget.addEventListener('change', () => schedulePreview());
element('retry-preview').addEventListener('click', () => schedulePreview(true));
element('clear-source').addEventListener('click', () => { sourceText.value = ''; schedulePreview(); sourceText.focus(); });

element('document-tab').addEventListener('click', () => { invalidatePreview(); showView(true); });
element('insert-translation').addEventListener('click', async () => {
  if (translating || signedOut || !previewReady || !previewRange) return;
  translating = true;
  updateDocumentActions();
  try {
    await previewRange.insert(translatedText.value, () => !signedOut);
    previewReady = false;
    await releasePreview();
    status('已替换原文，可使用 Word 撤销。继续翻译请重新选择范围。');
  } catch (error) { status((error as Error).message, true); }
  finally { translating = false; updateDocumentActions(); }
});
element('clear-translation-controls').addEventListener('click', async () => {
  if (translating || authorizing) return;
  translating = true;
  invalidatePreview();
  updateDocumentActions();
  element('restore-prompt').hidden = true;
  status('正在清除翻译控件…');
  try {
    await releasePreview();
    const count = await clearTranslationControls();
    status(count ? `已清除 ${count} 个翻译控件，文字和格式已保留。可重新选择内容翻译；这些范围无法再通过“恢复原文”还原。需要撤销清除时请使用 Word 撤销。请保存文档。`
      : '当前文档没有可清除的翻译控件。');
  } catch (error) { status((error as Error).message, true, true); }
  finally { translating = false; updateDocumentActions(); }
});
element('compact-backups').addEventListener('click', async () => {
  if (translating || authorizing) return;
  translating = true;
  updateDocumentActions();
  status('正在将文档内备份迁移到当前浏览器…');
  try {
    const migrated = await migrateDocumentBackups();
    status(migrated ? '原文备份已存入当前浏览器，文档内的旧备份已移除。请保存文档并重新下载检查大小。恢复需使用此浏览器，备份不会随文件共享。'
      : '此文档没有需要迁移的内嵌备份。新翻译的原文仅保存在当前浏览器。');
  } catch (error) { status((error as Error).message, true); }
  finally { translating = false; updateDocumentActions(); }
});
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
target.addEventListener('change', async () => {
  const fields = element<HTMLFieldSetElement>('settings-fields');
  fields.disabled = true;
  const previous = settings;
  const next = { target: target.value };
  settings = next;
  try {
    await saveSettings(next);
    updateDocumentActions();
  } catch (error) {
    settings = previous;
    target.value = previous.target;
    status((error as Error).message, true);
  }
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
  if (isWordOnline()) {
    element('document-summary').textContent = '将检查并备份原文，然后翻译当前文档。';
    element('document-detail').textContent = '完成后请保存文档；恢复原文需使用同一浏览器和插件地址。';
  }
  element('web-backup-notice').hidden = !isWordOnline();
  element('desktop-copy-notice').hidden = !isDesktopWord();
  if (isDesktopWord()) {
    element('compact-backups').hidden = true;
    element('backup-notice').textContent = '全文翻译每次创建新的副本。请保存副本，原文备份随副本保存；恢复请在副本中打开插件操作。恢复时保留范围外的编辑，文字已修改的译文会跳过。';
  }
  if (Office.context.document.addHandlerAsync) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      if (translating || signedOut || authorizing || element('selection-view').hidden) return;
      clearTimeout(previewTimer);
      invalidatePreview();
      previewTimer = setTimeout(() => void openPreview(activeScope, previewTarget.value), 250);
    });
  }
  settings = loadSettings();
  target.value = settings.target;
  previewTarget.value = settings.target;
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

element('dismiss-status').addEventListener('click', () => status(''));
element('copy-translation').addEventListener('click', async () => {
  if (!translatedText.value) { status('暂无可复制的译文。'); return; }
  try { await navigator.clipboard.writeText(translatedText.value); status('译文已复制'); }
  catch { status('复制失败，请选择译文后手动复制。', true); }
});
element('more-translation').addEventListener('click', () => {
  const menu = element('translation-menu');
  menu.hidden = !menu.hidden;
  element('more-translation').setAttribute('aria-expanded', String(!menu.hidden));
});
element('select-translation').addEventListener('click', () => {
  translatedText.focus(); translatedText.select();
  element('translation-menu').hidden = true;
  element('more-translation').setAttribute('aria-expanded', 'false');
});

import { api, authenticate } from '../shared/api';
import { hasLegacyBackup, rangeBackups, saveRangeBackups, migrateDocumentBackups, TAG_PREFIX, RangeBackup } from './backup';
import { prepareOoxml, rebaseOoxml } from './ooxml';
import { checkWebDocument, isWordOnline } from './web-safety';
import { stripNestedBackups } from './backup-ooxml';
import { COPY_KEY, desktopBackups, isDesktopWord, prepareDesktopCopy } from './desktop-copy';

export type Scope = 'selection' | 'paragraph' | 'body';
let busy = false;

function malformedOoxml(error: unknown): boolean {
  const failure = error as { code?: string; message?: string };
  return /ooxmlIsMalformed|ooxmlIsMalformated|InvalidOoxml/i.test(`${failure?.code || ''} ${failure?.message || ''}`);
}

async function syncStep(context: Word.RequestContext, step: string): Promise<void> {
  try { await context.sync(); }
  catch (error) {
    const failure = error as { code?: string; message?: string; debugInfo?: { errorLocation?: string } };
    const location = failure.debugInfo?.errorLocation;
    throw new Error(`${step}失败：${failure.code || failure.message || 'Word 未知错误'}${location ? `（${location}）` : ''}。`);
  }
}

function rangeFor(context: Word.RequestContext, scope: Scope): Word.Range {
  if (scope === 'body') return context.document.body.getRange();
  const selection = context.document.getSelection();
  return scope === 'paragraph' ? selection.paragraphs.getFirst().getRange() : selection;
}

async function ensureNoOverlap(context: Word.RequestContext, range: Word.Range, document: Word.Document | Word.DocumentCreated = context.document, readBackups = rangeBackups) {
  const controls = document.contentControls;
  controls.load('items/tag');
  await context.sync();
  // A failed operation can leave a wrapper without a committed checkpoint.
  // Remove only its wrapper, never its current contents or saved original.
  const pending = new Set((await readBackups()).filter(record => record.translatedText === undefined).map(record => record.tag));
  const failed = controls.items.filter(control => pending.has(control.tag));
  failed.forEach(control => control.delete(true));
  if (failed.length) await syncStep(context, '清理未完成的翻译标记');
  // Existing document/template controls can have a null tag at runtime.
  const relations = controls.items.filter(control => typeof control.tag === 'string'
    && control.tag.startsWith(TAG_PREFIX) && !pending.has(control.tag))
    .map(control => range.compareLocationWith(control.getRange()));
  await context.sync();
  const separate = ['Before', 'After', 'AdjacentBefore', 'AdjacentAfter', 'Unrelated'];
  if (relations.some(relation => !separate.includes(relation.value))) {
    throw new Error('所选范围包含尚未恢复的翻译，请先恢复该翻译，再重新选择内容。');
  }
}

export async function translateDocument(scope: Scope, target: string, shouldContinue = () => true): Promise<{ htmlBackup: boolean; skippedParagraphs?: number; desktopCopy?: boolean }> {
  if (busy) throw new Error('文档操作正在进行，请稍后重试。');
  busy = true;
  try {
    return await Word.run(async context => {
      const desktop = isDesktopWord() ? await prepareDesktopCopy(context, scope) : undefined;
      const document = desktop?.document || context.document;
      const storage = desktop ? desktopBackups(context, document) : undefined;
      const readBackups = storage?.read || rangeBackups;
      const saveBackups = storage?.save || saveRangeBackups;
      const range = desktop?.range || rangeFor(context, scope);
      context.trackedObjects.add(range);
      try {
        if (!desktop) await migrateDocumentBackups();
        await checkWebDocument(context);
        range.load('text');
        // Web selections/paragraphs also retain picture resources through OOXML.
        const nativeOoxml = scope === 'body' || isWordOnline() || !!desktop;
        const html = nativeOoxml ? undefined : range.getHtml();
        await syncStep(context, '读取翻译范围');
        let originalText = range.text;
        // Placeholder-only templates can expose empty Range.text. Let the
        // native parser determine whether their OOXML contains usable text.
        if (!originalText.trim() && !nativeOoxml) throw new Error(scope === 'selection' ? '请先选中需要翻译的文字。' : '当前范围没有可翻译的文字。');
        if (html && html.value.length > 1000000) throw new Error('文档内容过长，请选择较小范围分次翻译。');
        await ensureNoOverlap(context, range, document, readBackups);
        // Capture after retiring old pending wrappers so OOXML cannot revive them.
        const bodyOoxml = nativeOoxml ? range.getOoxml() : undefined;
        if (bodyOoxml) {
          range.load('text');
          await syncStep(context, '读取翻译范围原生结构（失败时停止，不转换为 HTML）');
          originalText = range.text;
        }
        const plan = bodyOoxml ? prepareOoxml(bodyOoxml.value) : undefined;
        const session = await authenticate();
        let translatedHtml = '';
        let translatedBody: { ooxml: string; skippedParagraphs: number } | undefined;
        let bodyTranslations: unknown;
        if (plan) {
          const result = await api<{ paragraphs: string[] }>('/api/translate/word', session.token, { paragraphs: plan.paragraphs, to: target });
          bodyTranslations = result.paragraphs;
        } else {
          const result = await api<{ html: string }>('/api/translate', session.token, { html: html!.value, to: target });
          if (typeof result.html !== 'string' || !result.html.trim() || result.html.length > 1000000) throw new Error('译文无效或超过显示限制。');
          translatedHtml = result.html;
        }
        if (!shouldContinue()) throw new Error('已取消翻译。');
        await checkWebDocument(context);
        await ensureNoOverlap(context, range, document, readBackups);
        const record: RangeBackup = { tag: TAG_PREFIX + crypto.randomUUID() };
        if (plan) {
          const latest = range.getOoxml();
          range.load('text');
          await syncStep(context, '读取全文备份');
          if (range.text !== originalText) throw new Error('翻译期间正文文字已更改，已取消替换，请重试。');
          // Back up the current layout, including formatting edits made while
          // waiting for the service. No stale package is used for replacement.
          translatedBody = rebaseOoxml(plan, latest.value, bodyTranslations);
          record.originalOoxml = latest.value;
        }
        else try {
          const original = range.getOoxml();
          await context.sync();
          if (!original.value.trim()) throw new Error('未能读取原文，已取消翻译。');
          record.originalOoxml = original.value;
        } catch (error) {
          if (!malformedOoxml(error)) throw error;
          if (!html!.value.trim()) throw new Error('Word 无法导出原文备份，已取消翻译。请保存并重新打开文档后重试。');
          // Some Word web hosts cannot export OOXML after an edit. Retain
          // the HTML captured from this exact tracked range before the request.
          // Do not fall back on insert errors or overwrite without a backup.
          record.originalHtml = html!.value;
        }
        const backups = await readBackups();
        // Persist the original before any replacement. An interrupted operation
        // leaves a pending record, which restoration never applies blindly.
        await saveBackups([...backups, record]);
        await checkWebDocument(context);
        range.load('text');
        const currentOoxml = bodyOoxml ? range.getOoxml() : undefined;
        await context.sync();
        if (!shouldContinue()) throw new Error('已取消翻译。');
        // Export packages are serialization results, not revision tokens.
        if (range.text !== originalText) {
          throw new Error('翻译期间原文已更改，已取消替换，请重试。');
        }
        if (plan && currentOoxml) translatedBody = rebaseOoxml(plan, currentOoxml.value, bodyTranslations);
        // Let Word import the content before wrapping it. A control created from
        // the original paragraph/selection can have incompatible boundaries.
        // Use the returned range, not the selection (which may have moved).
        const inserted = translatedBody
          ? range.insertOoxml(translatedBody.ooxml, Word.InsertLocation.replace)
          : range.insertHtml(translatedHtml, Word.InsertLocation.replace);
        await syncStep(context, '写入译文（原文备份已保留；如正文发生变化，请用 Word 撤销）');
        let control: Word.ContentControl | undefined;
        try {
          control = inserted.insertContentControl();
          await syncStep(context, '创建译文范围标记');
          control.tag = record.tag;
          control.title = '翻译内容（原文已备份）';
          await syncStep(context, '设置译文范围标记');
        } catch (error) {
          // Host batches are not transactions: creation may have succeeded
          // before a property write failed. Never delete the enclosed text.
          let cleanup = '';
          if (control) {
            try { control.delete(true); await context.sync(); }
            catch { cleanup = '标记清理也未完成。'; }
          }
          throw new Error(`${(error as Error).message}译文可能已写入，${cleanup}原文备份保留，请使用 Word 撤销本次操作。`);
        }
        const translated = control.getRange('Content');
        translated.load('text');
        await syncStep(context, '读取译文校验信息');
        record.translatedText = translated.text;
        try { await saveBackups([...backups, record]); }
        catch {
          throw new Error('译文已写入，但恢复校验信息保存失败。原文备份仍保留，请使用 Word 撤销本次翻译。');
        }
        return { htmlBackup: record.originalHtml !== undefined,
          ...(desktop ? { desktopCopy: true } : {}),
          ...(translatedBody?.skippedParagraphs ? { skippedParagraphs: translatedBody.skippedParagraphs } : {}) };
      } finally {
        try {
          context.trackedObjects.remove(range);
          await context.sync();
        } finally { await desktop?.finish(); }
      }
    });
  } finally { busy = false; }
}

export interface RestoreResult { restored: number; skipped: number; details?: string[]; cleaned?: number }
export async function restoreOriginalBody(): Promise<RestoreResult> {
  if (busy) throw new Error('文档操作正在进行，请稍后重试。');
  busy = true;
  try {
    const desktop = isDesktopWord();
    if (desktop && Office.context.document.settings.get(COPY_KEY) !== true) {
      throw new Error('请在翻译副本中打开插件并恢复原文。原始文档未进行翻译。');
    }
    return await Word.run(async context => {
    const backups = desktop ? await desktopBackups(context, context.document).read() : await rangeBackups();
    if (!backups.length) throw new Error(desktop ? '此副本没有可恢复的翻译备份。' : await hasLegacyBackup()
      ? '此文档仅有旧版整篇备份，无法只恢复翻译内容。为保留编辑，已停止整篇恢复。'
      : '当前浏览器没有此文档的原文备份。请使用翻译时的浏览器和插件地址；清除浏览器数据后无法恢复。');
      const result: RestoreResult = { restored: 0, skipped: 0 };
      const allControls = context.document.contentControls;
      allControls.load('items/tag');
      await context.sync();
      const known = new Set(backups.map(record => record.tag));
      const missing = allControls.items.filter(control => typeof control.tag === 'string'
        && control.tag.startsWith(TAG_PREFIX) && !known.has(control.tag));
      if (missing.length) {
        result.skipped = missing.length;
        result.details = ['部分翻译在当前浏览器中没有原文备份，请使用翻译时的浏览器恢复'];
      }
      // Keep records after restoration: Word Undo can bring the translated
      // control back, and its original must remain available in that case.
      for (const record of backups) {
        const controls = context.document.contentControls.getByTag(record.tag);
        controls.load('items');
        await context.sync();
        if (!controls.items.length) continue;
        if (record.translatedText === undefined) {
          // An incomplete operation has no trustworthy translated checkpoint.
          // Retire its misleading wrapper, preserving every character.
          controls.items.forEach(control => control.delete(true));
          await syncStep(context, '清理未完成的翻译标记');
          result.cleaned = (result.cleaned || 0) + controls.items.length;
          continue;
        }
        if (controls.items.length !== 1) {
          controls.items.forEach(control => { control.title = '翻译备份（标记重复，未恢复）'; });
          await context.sync();
          result.skipped++;
          (result.details ||= []).push('标记重复，无法唯一定位原文');
          continue;
        }
        const control = controls.items[0];
        const range = control.getRange('Content');
        range.load('text');
        await context.sync();
        // Word HTML exports aren't stable across saves/reopens. Compare exact
        // text, including whitespace. Restoration also restores the original
        // formatting within this control, as stated in the confirmation UI.
        if (range.text !== record.translatedText) {
          control.title = '翻译备份（文字不匹配，未恢复）';
          await context.sync();
          result.skipped++;
          (result.details ||= []).push('当前文字与保存的译文不一致（可能已编辑或撤销），未覆盖');
          continue;
        }
        if (record.originalOoxml) control.insertOoxml(stripNestedBackups(record.originalOoxml), Word.InsertLocation.replace);
        else control.insertHtml(record.originalHtml!, Word.InsertLocation.replace);
        await syncStep(context, '恢复翻译范围原文');
        control.delete(true);
        await context.sync();
        result.restored++;
      }
      return result;
    });
  } finally { busy = false; }
}

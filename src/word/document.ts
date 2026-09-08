import { api, authenticate } from '../shared/api';
import { hasLegacyBackup, rangeBackups, saveRangeBackups, TAG_PREFIX, RangeBackup } from './backup';

export type Scope = 'selection' | 'paragraph' | 'body';
let busy = false;

function rangeFor(context: Word.RequestContext, scope: Scope): Word.Range {
  if (scope === 'body') return context.document.body.getRange();
  const selection = context.document.getSelection();
  return scope === 'paragraph' ? selection.paragraphs.getFirst().getRange() : selection;
}

async function ensureNoOverlap(context: Word.RequestContext, range: Word.Range) {
  const controls = context.document.contentControls;
  controls.load('items/tag');
  await context.sync();
  const relations = controls.items.filter(control => control.tag.startsWith(TAG_PREFIX))
    .map(control => range.compareLocationWith(control.getRange()));
  await context.sync();
  const separate = ['Before', 'After', 'AdjacentBefore', 'AdjacentAfter', 'Unrelated'];
  if (relations.some(relation => !separate.includes(relation.value))) {
    throw new Error('所选范围包含尚未恢复的翻译，请先恢复该翻译，再重新选择内容。');
  }
}

export async function translateDocument(scope: Scope, target: string, shouldContinue = () => true): Promise<void> {
  if (busy) throw new Error('文档操作正在进行，请稍后重试。');
  busy = true;
  try {
    await Word.run(async context => {
      const range = rangeFor(context, scope);
      context.trackedObjects.add(range);
      try {
        range.load('text');
        const html = range.getHtml();
        await context.sync();
        const originalText = range.text;
        if (!originalText.trim()) throw new Error(scope === 'selection' ? '请先选中需要翻译的文字。' : '当前范围没有可翻译的文字。');
        if (html.value.length > 1000000) throw new Error('文档内容过长，请选择较小范围分次翻译。');
        await ensureNoOverlap(context, range);
        const session = await authenticate();
        const result = await api<{ html: string }>('/api/translate', session.token, { html: html.value, to: target });
        if (typeof result.html !== 'string' || !result.html.trim() || result.html.length > 1000000) throw new Error('译文无效或超过显示限制。');
        if (!shouldContinue()) throw new Error('已取消翻译。');
        await ensureNoOverlap(context, range);
        const original = range.getOoxml();
        await context.sync();
        if (!original.value.trim()) throw new Error('未能读取原文，已取消翻译。');
        const backups = rangeBackups();
        const record: RangeBackup = { tag: TAG_PREFIX + crypto.randomUUID(), originalOoxml: original.value };
        // Persist the original before any replacement. An interrupted operation
        // leaves a pending record, which restoration never applies blindly.
        await saveRangeBackups([...backups, record]);
        range.load('text');
        await context.sync();
        if (!shouldContinue()) throw new Error('已取消翻译。');
        // HTML exports are serialization results, not revision tokens. Word
        // can change export metadata when settings are saved. Compare exact
        // text here; formatting-only edits during the request aren't detected.
        if (range.text !== originalText) {
          throw new Error('翻译期间原文已更改，已取消替换，请重试。');
        }
        const control = range.insertContentControl();
        control.tag = record.tag;
        control.title = '翻译内容（可恢复原文）';
        control.insertHtml(result.html, Word.InsertLocation.replace);
        const translated = control.getRange('Content');
        translated.load('text');
        const translatedHtml = translated.getHtml();
        await context.sync();
        record.translatedHtml = translatedHtml.value;
        record.translatedText = translated.text;
        try { await saveRangeBackups([...backups, record]); }
        catch {
          throw new Error('译文已写入，但恢复校验信息保存失败。原文备份仍保留，请使用 Word 撤销本次翻译。');
        }
      } finally {
        context.trackedObjects.remove(range);
        await context.sync();
      }
    });
  } finally { busy = false; }
}

export interface RestoreResult { restored: number; skipped: number }
export async function restoreOriginalBody(): Promise<RestoreResult> {
  if (busy) throw new Error('文档操作正在进行，请稍后重试。');
  busy = true;
  try {
    const backups = rangeBackups();
    if (!backups.length) throw new Error(hasLegacyBackup()
      ? '此文档仅有旧版整篇备份，无法只恢复翻译内容。为保留编辑，已停止整篇恢复。'
      : '当前文档没有可恢复的翻译范围备份。');
    return await Word.run(async context => {
      const result: RestoreResult = { restored: 0, skipped: 0 };
      // Keep records after restoration: Word Undo can bring the translated
      // control back, and its original must remain available in that case.
      for (const record of backups) {
        const controls = context.document.contentControls.getByTag(record.tag);
        controls.load('items');
        await context.sync();
        if (controls.items.length !== 1 || record.translatedHtml === undefined || record.translatedText === undefined) {
          if (controls.items.length) result.skipped++;
          continue;
        }
        const control = controls.items[0];
        const range = control.getRange('Content');
        range.load('text');
        const html = range.getHtml();
        await context.sync();
        // Fail closed on formatting/serialization differences too. A false
        // conflict is preferable to silently overwriting a user's edits.
        if (range.text !== record.translatedText || html.value !== record.translatedHtml) {
          result.skipped++;
          continue;
        }
        control.insertOoxml(record.originalOoxml, Word.InsertLocation.replace);
        control.delete(true);
        await context.sync();
        result.restored++;
      }
      return result;
    });
  } finally { busy = false; }
}

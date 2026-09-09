import type { Scope } from './document';
import type { RangeBackup } from './backup';
import { DOCUMENT_ID_KEY } from './backup';
import { accessLocalBackup } from './backup-store';

export const COPY_KEY = 'wordTranslation.desktopCopy.v1';
const BACKUPS_KEY = 'wordTranslation.ranges.v2';

export function isDesktopWord(): boolean {
  return Office.context.platform !== undefined && (Office.context.platform === Office.PlatformType.PC || Office.context.platform === Office.PlatformType.Mac);
}

function result<T>(start: (callback: (value: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => start(value => value.status === Office.AsyncResultStatus.Succeeded
    ? resolve(value.value) : reject(new Error(value.error?.message || '读取文档副本失败。'))));
}

export async function exportDocument(): Promise<string> {
  const file = await result<Office.File>(cb => Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, cb));
  try {
    const chunks: string[] = [];
    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await result<Office.Slice>(cb => file.getSliceAsync(index, cb));
      const bytes = slice.data as number[];
      // Encode once after joining: slice boundaries need not align to Base64 triplets.
      for (let offset = 0; offset < bytes.length; offset += 8192) {
        chunks.push(String.fromCharCode(...bytes.slice(offset, offset + 8192)));
      }
    }
    return btoa(chunks.join(''));
  } finally { await result<void>(cb => file.closeAsync(cb)); }
}

export function desktopBackups(context: Word.RequestContext, doc: Word.Document | Word.DocumentCreated) {
  return {
    async read(): Promise<RangeBackup[]> {
      const setting = doc.settings.getItemOrNullObject(BACKUPS_KEY);
      setting.load('value');
      await context.sync();
      if (setting.isNullObject) return [];
      const records = setting.value as RangeBackup[];
      if (!Array.isArray(records) || records.some(r => !r || typeof r.tag !== 'string'
        || !(typeof r.originalOoxml === 'string' || typeof r.originalHtml === 'string')
        || (r.translatedText !== undefined && typeof r.translatedText !== 'string'))) {
        throw new Error('副本原文备份无效，已停止操作。');
      }
      return records.map(r => ({ ...r }));
    },
    async save(records: RangeBackup[]): Promise<void> {
      doc.settings.add(BACKUPS_KEY, records);
      await context.sync();
    },
  };
}

export async function prepareDesktopCopy(context: Word.RequestContext, scope: Scope) {
  const source = context.document;
  const selection = scope === 'body' ? source.body.getRange() : source.getSelection();
  const range = scope === 'paragraph' ? selection.paragraphs.getFirst().getRange() : selection;
  range.load('text');
  await context.sync();
  if (!range.text.trim()) throw new Error(scope === 'selection' ? '请先选中需要翻译的文字。' : '当前范围没有可翻译的文字。');
  if (Office.context.document.settings.get(COPY_KEY) === true) {
    return { document: source as Word.Document | Word.DocumentCreated, range, created: false, finish: async () => {} };
  }
  if (!Office.context.requirements.isSetSupported('WordApiHiddenDocument', '1.4')
    || !Office.context.requirements.isSetSupported('WordApi', '1.4')) {
    throw new Error('当前 Word 版本不支持安全创建翻译副本，请更新桌面 Word 后重试。');
  }
  // A unique temporary bookmark carries exact range boundaries through the full
  // DOCX export, including repeated text and selections spanning paragraphs.
  const bookmark = 'wt' + crypto.randomUUID().replace(/-/g, '');
  let base64: string;
  try {
    if (scope !== 'body') { range.insertBookmark(bookmark); await context.sync(); }
    base64 = await exportDocument();
  } finally {
    if (scope !== 'body') { source.deleteBookmark(bookmark); await context.sync(); }
  }
  const copy = context.application.createDocument(base64);
  await context.sync();
  try {
    copy.settings.add(COPY_KEY, true);
    // A copied browser-backup ID must never make this document share future
    // backup writes with its source. Bring existing originals into the copy.
    const sourceId = Office.context.document.settings.get(DOCUMENT_ID_KEY);
    if (sourceId) {
      const local = await accessLocalBackup<{ ranges: RangeBackup[] }>(sourceId);
      const embedded = await desktopBackups(context, copy).read();
      const records = new Map((local?.ranges || []).map(record => [record.tag, record]));
      embedded.forEach(record => records.set(record.tag, record));
      await desktopBackups(context, copy).save([...records.values()]);
    }
    copy.settings.add(DOCUMENT_ID_KEY, crypto.randomUUID());
    const copiedRange = scope === 'body' ? copy.body.getRange() : copy.getBookmarkRange(bookmark);
    context.trackedObjects.add(copiedRange);
    copiedRange.load('text');
    await context.sync();
    if (scope !== 'body') { copy.deleteBookmark(bookmark); await context.sync(); }
    if (copiedRange.text !== range.text) throw new Error('副本中的翻译范围与原文不一致，已停止翻译。');
    return { document: copy as Word.Document | Word.DocumentCreated, range: copiedRange, created: true,
      finish: async () => { copy.open(); await context.sync(); } };
  } catch (error) {
    copy.open(); await context.sync();
    throw error;
  }
}

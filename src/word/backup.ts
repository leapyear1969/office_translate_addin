import { stripNestedBackups } from './backup-ooxml';
import { accessLocalBackup } from './backup-store';

const KEY = 'wordTranslation.ranges.v2';
const LEGACY_KEY = 'wordTranslation.originalBody.v1';
export const DOCUMENT_ID_KEY = 'wordTranslation.localBackupId.v1';
export const TAG_PREFIX = 'wordTranslation:';

export interface RangeBackup {
  tag: string;
  originalOoxml?: string;
  originalHtml?: string;
  // Missing until the translated content has been read back and persisted.
  translatedHtml?: string;
  translatedText?: string;
}
interface LocalBackup { ranges: RangeBackup[]; legacy?: unknown }

function validate(value: unknown): RangeBackup[] {
  if (value == null) return [];
  if (!Array.isArray(value) || value.some(item => !item || typeof item.tag !== 'string'
    || !item.tag.startsWith(TAG_PREFIX)
    || !((typeof item.originalOoxml === 'string' && item.originalOoxml.trim())
      || (typeof item.originalHtml === 'string' && item.originalHtml.trim()))
    || (item.originalOoxml !== undefined && typeof item.originalOoxml !== 'string')
    || (item.originalHtml !== undefined && typeof item.originalHtml !== 'string')
    || (item.translatedHtml !== undefined && typeof item.translatedHtml !== 'string')
    || (item.translatedText !== undefined && typeof item.translatedText !== 'string'))
    || new Set(value.map(item => item.tag)).size !== value.length) {
    throw new Error('翻译范围备份无效，已取消操作。');
  }
  return value.map(item => ({ ...item }));
}

function documentId(): string | undefined {
  const id: unknown = Office.context.document.settings.get(DOCUMENT_ID_KEY);
  if (id == null) return undefined;
  if (typeof id !== 'string' || !/^[0-9a-f-]{36}$/i.test(id)) throw new Error('文档本地备份标识无效，已取消操作。');
  return id;
}

function compact(backups: RangeBackup[]): RangeBackup[] {
  return validate(backups).map(item => ({ ...item,
    ...(item.originalOoxml ? { originalOoxml: stripNestedBackups(item.originalOoxml) } : {}),
  }));
}

async function write(id: string, ranges: RangeBackup[], legacy?: unknown, migrating = false): Promise<void> {
  await accessLocalBackup<LocalBackup>(id, current => {
    const records = new Map(compact(current?.ranges || []).map(item => [item.tag, item]));
    for (const record of compact(ranges)) {
      const previous = records.get(record.tag);
      if (previous && (previous.originalOoxml !== record.originalOoxml || previous.originalHtml !== record.originalHtml)) {
        throw new Error('本地备份标记冲突，已停止保存。');
      }
      // A stale export or parallel writer must not erase committed checkpoints
      // or another range's original. Keep all history for Word Undo.
      records.set(record.tag, previous && (record.translatedText === undefined
        || (migrating && previous.translatedText !== undefined))
        ? previous : record);
    }
    return { ...current, ranges: [...records.values()],
      ...(legacy != null && current?.legacy == null ? { legacy } : {}) };
  });
}

async function saveReference(id: string, removeEmbedded: boolean): Promise<void> {
  const storage = Office.context.document.settings;
  const keys = removeEmbedded ? [DOCUMENT_ID_KEY, KEY, LEGACY_KEY] : [DOCUMENT_ID_KEY];
  const previous = keys.map(key => storage.get(key));
  try {
    storage.set(DOCUMENT_ID_KEY, id);
    if (removeEmbedded) { storage.remove(KEY); storage.remove(LEGACY_KEY); }
    await new Promise<void>((resolve, reject) => storage.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error('无法保存文档备份标识。'));
    }));
  } catch (error) {
    keys.forEach((key, index) => {
      if (previous[index] == null) storage.remove(key);
      else storage.set(key, previous[index]);
    });
    throw error;
  }
}

/** Commit locally before removing either generation of embedded backup.
 * A failed settings save leaves the source intact; retry merges idempotently. */
export async function migrateDocumentBackups(): Promise<boolean> {
  const storage = Office.context.document.settings;
  const embedded: unknown = storage.get(KEY);
  const legacy: unknown = storage.get(LEGACY_KEY);
  if (embedded == null && legacy == null) return false;
  const ranges = compact(validate(embedded));
  const id = documentId() || crypto.randomUUID();
  try {
    await write(id, ranges, legacy, true);
    await saveReference(id, true);
    return true;
  } catch (error) {
    throw new Error('迁移翻译备份失败，文档内原备份已保留。' + (error as Error).message);
  }
}

async function read(): Promise<LocalBackup | undefined> {
  await migrateDocumentBackups();
  const id = documentId();
  if (!id) return undefined;
  try { return await accessLocalBackup<LocalBackup>(id); }
  catch { throw new Error('无法读取浏览器本地备份，请检查浏览器存储权限并重试。'); }
}

export async function rangeBackups(): Promise<RangeBackup[]> {
  return validate((await read())?.ranges);
}

export async function hasLegacyBackup(): Promise<boolean> {
  return (await read())?.legacy != null;
}

export async function saveRangeBackups(backups: RangeBackup[]): Promise<void> {
  await migrateDocumentBackups();
  const existing = documentId();
  const id = existing || crypto.randomUUID();
  try {
    await write(id, backups);
    // Only a short ID is persisted in the document, never backup content.
    if (!existing) await saveReference(id, false);
  } catch (error) {
    throw new Error('保存翻译范围备份失败，未能提交浏览器本地备份或文档标识。请检查浏览器存储权限和可用空间。' + (error as Error).message);
  }
}

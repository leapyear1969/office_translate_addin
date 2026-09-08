const KEY = 'wordTranslation.ranges.v2';
const LEGACY_KEY = 'wordTranslation.originalBody.v1';
export const TAG_PREFIX = 'wordTranslation:';

export interface RangeBackup {
  tag: string;
  originalOoxml: string;
  // Missing until the translated content has been read back and persisted.
  translatedHtml?: string;
  translatedText?: string;
}

export function rangeBackups(): RangeBackup[] {
  const value: unknown = Office.context.document.settings.get(KEY);
  if (value == null) return [];
  if (!Array.isArray(value) || value.some(item => !item || typeof item.tag !== 'string'
    || !item.tag.startsWith(TAG_PREFIX) || typeof item.originalOoxml !== 'string' || !item.originalOoxml.trim()
    || (item.translatedHtml !== undefined && typeof item.translatedHtml !== 'string')
    || (item.translatedText !== undefined && typeof item.translatedText !== 'string'))
    || new Set(value.map(item => item.tag)).size !== value.length) {
    throw new Error('翻译范围备份无效，已取消操作。');
  }
  return value.map(item => ({ ...item }));
}

export function hasLegacyBackup(): boolean {
  return Office.context.document.settings.get(LEGACY_KEY) != null;
}

export async function saveRangeBackups(backups: RangeBackup[]): Promise<void> {
  const storage = Office.context.document.settings;
  const previous = storage.get(KEY);
  try {
    storage.set(KEY, backups.map(item => ({ ...item })));
    await new Promise<void>((resolve, reject) => storage.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error('保存失败'));
    }));
  } catch {
    if (previous == null) storage.remove(KEY);
    else storage.set(KEY, previous);
    throw new Error('保存翻译范围备份失败，请重试。');
  }
}

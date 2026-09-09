const WEB_EXTENSION = 'http://schemas.microsoft.com/office/webextensions/webextension/2010/11';
const BACKUP_KEYS = new Set(['wordTranslation.ranges.v2', 'wordTranslation.originalBody.v1', 'wordTranslation.localBackupId.v1']);

/** Do not nest our history or re-import stale history with range content.
 * Preserve other settings, add-ins, relationships and all content parts. */
export function stripNestedBackups(xml: string): string {
  if (![...BACKUP_KEYS].some(key => xml.includes(key))) return xml;
  if (xml.length > 50000000 || /<!DOCTYPE/i.test(xml)) throw new Error('原文备份结构过大或不受支持，无法安全精简。');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('原文备份结构无效，无法安全精简。');
  let changed = false;
  for (const property of Array.from(doc.getElementsByTagNameNS(WEB_EXTENSION, 'property'))) {
    if (BACKUP_KEYS.has(property.getAttribute('name') || '')) {
      property.remove();
      changed = true;
    }
  }
  return changed ? new XMLSerializer().serializeToString(doc) : xml;
}

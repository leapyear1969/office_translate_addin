/** @jest-environment jsdom */
import { migrateDocumentBackups, rangeBackups, saveRangeBackups, hasLegacyBackup, DOCUMENT_ID_KEY, TAG_PREFIX } from './backup';
import { stripNestedBackups } from './backup-ooxml';
import { prepareOoxml } from './ooxml';
import { wordPackage } from './test-fixtures/package';
import { resetIndexedDB } from './test-fixtures/indexeddb';
import { IDBObjectStore } from 'fake-indexeddb';
import { accessLocalBackup } from './backup-store';

const key = 'wordTranslation.ranges.v2';
const nested = () => wordPackage().replace('</pkg:package>', `<pkg:part pkg:name="/word/webextensions/webextension.xml"><pkg:xmlData><we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:properties><we:property name="${key}" value="${'nested-history'.repeat(1000)}"/><we:property name="wordTranslation.originalBody.v1" value="old"/><we:property name="wordTranslation.preferences.v1" value="ja"/><we:property name="other.addin" value="keep"/></we:properties></we:webextension></pkg:xmlData></pkg:part></pkg:package>`);
function setup(initial?: unknown) {
  jest.restoreAllMocks();
  resetIndexedDB();
  const values = new Map<string, unknown>([[key, initial]]);
  const storage = { get: (name: string) => values.get(name),
    set: jest.fn((name: string, value: unknown) => values.set(name, value)),
    remove: jest.fn((name: string) => values.delete(name)),
    saveAsync: jest.fn((cb: Function) => cb({ status: 'succeeded' })) };
  (globalThis as any).Office = { context: { document: { settings: storage } }, AsyncResultStatus: { Succeeded: 'succeeded' } };
  return storage;
}

test('removes only nested backup settings and preserves body, images, preferences and other add-ins', () => {
  const xml = stripNestedBackups(nested());
  expect(xml).not.toContain('nested-history');
  expect(xml).not.toContain('wordTranslation.originalBody.v1');
  expect(xml).toContain('wordTranslation.preferences.v1');
  expect(xml).toContain('other.addin');
  expect(xml).toContain('aW1hZ2U=');
  expect(xml).toContain('Original');
  expect(stripNestedBackups(xml)).toBe(xml);
  expect(prepareOoxml(nested()).apply(['<p><span id="r0">译文</span></p>']).ooxml).not.toContain('nested-history');
});

const record = (name = 'one') => ({ tag: TAG_PREFIX + name, originalOoxml: '<original/>', translatedText: '译文' });

test('large originals are stored in IndexedDB; document settings contain only a short identifier', async () => {
  const storage = setup();
  const large = { tag: TAG_PREFIX + 'large', originalHtml: '<p>' + '中&'.repeat(600000) + '</p>' };
  await saveRangeBackups([large]);
  const id = storage.get(DOCUMENT_ID_KEY) as string;
  expect(id).toHaveLength(36);
  expect(storage.set).toHaveBeenCalledTimes(1);
  expect(storage.set).toHaveBeenCalledWith(DOCUMENT_ID_KEY, id);
  expect(storage.get(key)).toBeUndefined();
  expect(await accessLocalBackup(id)).toEqual({ ranges: [large] });
  expect(await rangeBackups()).toEqual([large]);
  await saveRangeBackups([{ ...large, translatedText: '翻译' }]);
  expect(storage.saveAsync).toHaveBeenCalledTimes(1);
});

test('migration commits local originals before removing both embedded backup versions', async () => {
  const original = [{ ...record(), originalOoxml: nested() }];
  const storage = setup(original);
  storage.set('wordTranslation.originalBody.v1', '<legacy-body/>');
  storage.set('wordTranslation.preferences.v1', { target: 'ja' });
  expect(await migrateDocumentBackups()).toBe(true);
  expect(await rangeBackups()).toEqual([{ ...original[0], originalOoxml: stripNestedBackups(nested()) }]);
  expect(await hasLegacyBackup()).toBe(true);
  expect(storage.get(key)).toBeUndefined();
  expect(storage.get('wordTranslation.originalBody.v1')).toBeUndefined();
  expect(storage.get('wordTranslation.preferences.v1')).toEqual({ target: 'ja' });
  expect(await accessLocalBackup(storage.get(DOCUMENT_ID_KEY) as string)).toMatchObject({ legacy: '<legacy-body/>' });
  expect(await migrateDocumentBackups()).toBe(false);
  expect(storage.saveAsync).toHaveBeenCalledTimes(1);
});

test('an aborted local transaction never removes embedded originals even if put succeeded', async () => {
  const original = [record()];
  const storage = setup(original);
  const put = IDBObjectStore.prototype.put;
  jest.spyOn(IDBObjectStore.prototype, 'put').mockImplementationOnce(function (this: IDBObjectStore, value, id) {
    const request = put.call(this, value, id);
    request.onsuccess = () => this.transaction.abort();
    return request;
  });
  await expect(migrateDocumentBackups()).rejects.toThrow('原备份已保留');
  expect(storage.get(key)).toEqual(original);
  expect(storage.remove).not.toHaveBeenCalled();
  expect(storage.saveAsync).not.toHaveBeenCalled();
});

test('failed document cleanup retains embedded data and can be retried', async () => {
  const storage = setup([record()]);
  storage.saveAsync.mockImplementationOnce(cb => cb({ status: 'failed' }));
  await expect(migrateDocumentBackups()).rejects.toThrow('原备份已保留');
  expect(storage.get(key)).toEqual([record()]);
  expect(storage.get(DOCUMENT_ID_KEY)).toBeUndefined();
  expect(await migrateDocumentBackups()).toBe(true);
  expect(await rangeBackups()).toEqual([record()]);
});

test('unavailable or quota-limited browser storage does not fall back to document data', async () => {
  const storage = setup();
  jest.spyOn(IDBObjectStore.prototype, 'put').mockImplementationOnce(() => { throw new DOMException('Full', 'QuotaExceededError'); });
  await expect(saveRangeBackups([record()])).rejects.toThrow('可用空间');
  expect(storage.set).not.toHaveBeenCalled();
  Object.defineProperty(globalThis, 'indexedDB', { configurable: true, value: undefined });
  await expect(saveRangeBackups([record()])).rejects.toThrow('IndexedDB');
  expect(storage.set).not.toHaveBeenCalled();
});

test('database contents survive module reload; another document and another browser cannot see them', async () => {
  const storage = setup();
  await saveRangeBackups([record()]);
  const id = storage.get(DOCUMENT_ID_KEY);
  jest.resetModules();
  const reopened: typeof import('./backup') = require('./backup');
  expect(await reopened.rangeBackups()).toEqual([record()]);
  storage.remove(DOCUMENT_ID_KEY);
  expect(await reopened.rangeBackups()).toEqual([]);
  await reopened.saveRangeBackups([record('two')]);
  expect(storage.get(DOCUMENT_ID_KEY)).not.toBe(id);
  expect(await reopened.rangeBackups()).toEqual([record('two')]);
  storage.set(DOCUMENT_ID_KEY, id);
  expect(await reopened.rangeBackups()).toEqual([record()]);
  resetIndexedDB();
  expect(await reopened.rangeBackups()).toEqual([]);
});

test('concurrent writes for an established document merge ranges and stale pending saves keep checkpoints', async () => {
  setup();
  await saveRangeBackups([record()]);
  await Promise.all([saveRangeBackups([record('two')]), saveRangeBackups([record('three')])]);
  await saveRangeBackups([{ tag: record().tag, originalOoxml: '<original/>' }]);
  const records = await rangeBackups();
  expect(records).toHaveLength(3);
  expect(records.find(item => item.tag === record().tag)).toEqual(record());
});

test('a conflicting original aborts the entire merge and preserves existing records', async () => {
  setup();
  await saveRangeBackups([record()]);
  await expect(saveRangeBackups([record('two'), { ...record(), originalOoxml: '<wrong/>' }])).rejects.toThrow('冲突');
  expect(await rangeBackups()).toEqual([record()]);
});

test('invalid embedded backups are retained without starting migration', async () => {
  const storage = setup([{ tag: 'wrong' }]);
  await expect(migrateDocumentBackups()).rejects.toThrow('备份无效');
  expect(storage.remove).not.toHaveBeenCalled();
});

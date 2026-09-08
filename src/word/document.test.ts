import { translateDocument, restoreOriginalBody } from './document';
import { api, authenticate } from '../shared/api';
import { rangeBackups, saveRangeBackups, TAG_PREFIX } from './backup';
import { randomUUID } from 'crypto';
jest.mock('../shared/api');

function setup() {
  (globalThis as any).crypto = { randomUUID };
  const values = new Map<string, unknown>();
  const storage = {
    get: jest.fn((key: string) => values.get(key)),
    set: jest.fn((key: string, value: unknown) => values.set(key, value)),
    remove: jest.fn((key: string) => values.delete(key)),
    saveAsync: jest.fn((cb: Function) => cb({ status: 'succeeded' })),
  };
  (globalThis as any).Office = { context: { document: { settings: storage } }, AsyncResultStatus: { Succeeded: 'succeeded' } };
  const controls: any[] = [];
  const makeRange = () => {
    const range: any = {
      text: 'Original', html: '<p>Original</p>', ooxml: '<original/>', relation: 'Before',
      load: jest.fn(), getHtml: jest.fn(() => ({ value: range.html })),
      getOoxml: jest.fn(() => ({ value: range.ooxml })),
      compareLocationWith: jest.fn(() => ({ value: range.relation })),
      insertContentControl: jest.fn(() => {
        const control: any = {
          tag: '', title: '', getRange: jest.fn(() => range),
          insertHtml: jest.fn((html: string) => { range.html = html; range.text = '译文'; }),
          insertOoxml: jest.fn((xml: string) => { range.ooxml = xml; range.text = 'Original'; range.html = '<p>Original</p>'; }),
          delete: jest.fn(() => { controls.splice(controls.indexOf(control), 1); }),
        };
        controls.push(control);
        return control;
      }),
    };
    return range;
  };
  const paragraph = makeRange();
  const body = makeRange();
  const selection = makeRange();
  selection.paragraphs = { getFirst: () => ({ getRange: () => paragraph }) };
  const collection = { items: controls, load: jest.fn(), getByTag: (tag: string) => ({
    items: controls.filter(control => control.tag === tag), load: jest.fn(),
  }) };
  const context = {
    document: { contentControls: collection, body: { getRange: () => body, insertOoxml: jest.fn() }, getSelection: jest.fn(() => selection) },
    sync: jest.fn(async () => {}), trackedObjects: { add: jest.fn(), remove: jest.fn() },
  };
  (globalThis as any).Word = { run: (fn: Function) => fn(context), InsertLocation: { replace: 'Replace' } };
  jest.mocked(authenticate).mockResolvedValue({ token: 'token', user: {} as any });
  jest.mocked(api).mockResolvedValue({ html: '<p>译文</p>' });
  return { context, selection, paragraph, body, storage, controls };
}

test.each(['selection', 'paragraph', 'body'] as const)('backs up and translates only the %s range', async scope => {
  const ranges = setup();
  ranges[scope].ooxml = `<${scope}/>`;
  await translateDocument(scope, 'ja');
  expect(rangeBackups()[0].originalOoxml).toBe(`<${scope}/>`);
  expect(rangeBackups()[0].translatedText).toBe('译文');
  expect(ranges.controls[0].insertHtml).toHaveBeenCalledWith('<p>译文</p>', 'Replace');
  expect(ranges.context.document.body.insertOoxml).not.toHaveBeenCalled();
  expect(ranges.context.trackedObjects.remove).toHaveBeenCalledWith(ranges[scope]);
});

test('restores translation while preserving edits in other ranges, without authentication', async () => {
  const { selection, body, controls } = setup();
  await translateDocument('selection', 'ja');
  const control = controls[0];
  body.text = 'Later saved edits elsewhere';
  jest.mocked(authenticate).mockClear(); jest.mocked(api).mockClear();
  expect(await restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
  expect(selection.text).toBe('Original');
  expect(body.text).toBe('Later saved edits elsewhere');
  expect(control.delete).toHaveBeenCalledWith(true);
  expect(authenticate).not.toHaveBeenCalled();
  expect(api).not.toHaveBeenCalled();
  expect(await restoreOriginalBody()).toEqual({ restored: 0, skipped: 0 });
});

test.each(['User correction', '', '译文 '])('skips edited translation text %j and restores other independent translations', async edited => {
  const { selection, paragraph, controls } = setup();
  await translateDocument('selection', 'ja');
  await translateDocument('paragraph', 'ja');
  selection.text = edited;
  const conflicted = controls[0];
  expect(await restoreOriginalBody()).toEqual({ restored: 1, skipped: 1 });
  expect(conflicted.insertOoxml).not.toHaveBeenCalled();
  expect(paragraph.text).toBe('Original');
  expect(selection.text).toBe(edited);
});

test.each(['<p id="new-export"><span>译文</span></p>', '<p><b>译文</b></p>'])('restores original text and formatting when only HTML changes: %s', async html => {
  const { selection, body, controls } = setup();
  await translateDocument('selection', 'ja');
  const control = controls[0];
  selection.html = html;
  body.text = 'Hello, I am Jason';
  expect(await restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
  expect(control.insertOoxml).toHaveBeenCalledWith('<original/>', 'Replace');
  expect(body.text).toBe('Hello, I am Jason');
});

test('existing backups with a stale translated HTML snapshot remain restorable', async () => {
  const { selection, storage } = setup();
  await translateDocument('selection', 'ja');
  const records = rangeBackups();
  records[0].translatedHtml = '<p id="old-export">译文</p>';
  storage.set('wordTranslation.ranges.v2', records);
  selection.html = '<p id="new-export">译文</p>';
  expect(await restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
});

test('new translation after restoration backs up the latest edited original', async () => {
  const { selection, controls } = setup();
  await translateDocument('selection', 'ja');
  await restoreOriginalBody();
  selection.text = 'Updated original'; selection.html = '<p>Updated original</p>'; selection.ooxml = '<updated/>';
  await translateDocument('selection', 'ja');
  const control = controls[0];
  await restoreOriginalBody();
  expect(control.insertOoxml).toHaveBeenCalledWith('<updated/>', 'Replace');
});

test('persisted records restore after module reload and retain originals for Word undo', async () => {
  const { controls } = setup();
  await translateDocument('selection', 'ja');
  const control = controls[0];
  jest.resetModules();
  const reloaded = require('./document');
  expect(await reloaded.restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
  control.insertHtml('<p>译文</p>'); controls.push(control);
  expect(await reloaded.restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
});

test.each(['Equal', 'Inside', 'Contains', 'OverlapsBefore', 'OverlapsAfter'])('rejects overlapping translation: %s', async relation => {
  const { selection } = setup();
  await translateDocument('selection', 'ja');
  selection.relation = relation;
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('尚未恢复');
  expect(selection.insertContentControl).toHaveBeenCalledTimes(1);
});

test('deleted controls never cause fallback to whole-document restoration', async () => {
  const { controls, context } = setup();
  await translateDocument('selection', 'ja');
  controls.splice(0);
  expect(await restoreOriginalBody()).toEqual({ restored: 0, skipped: 0 });
  expect(context.document.body.insertOoxml).not.toHaveBeenCalled();
});

test('duplicate tags are skipped rather than restoring the wrong copy', async () => {
  const { controls } = setup();
  await translateDocument('selection', 'ja');
  controls.push(controls[0]);
  expect(await restoreOriginalBody()).toEqual({ restored: 0, skipped: 1 });
  expect(controls[0].insertOoxml).not.toHaveBeenCalled();
});

test('legacy backup cannot overwrite later edits', async () => {
  const { storage, context } = setup();
  storage.set('wordTranslation.originalBody.v1', '<old-body/>');
  await expect(restoreOriginalBody()).rejects.toThrow('旧版整篇备份');
  expect(context.document.body.insertOoxml).not.toHaveBeenCalled();
});

test('backup persistence failure prevents replacement and permits retry', async () => {
  const { storage, selection } = setup();
  storage.saveAsync.mockImplementationOnce(cb => cb({ status: 'failed' }));
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('保存翻译范围备份失败');
  expect(selection.insertContentControl).not.toHaveBeenCalled();
  expect(rangeBackups()).toEqual([]);
  await translateDocument('selection', 'ja');
  expect(selection.insertContentControl).toHaveBeenCalledTimes(1);
});

test('checkpoint failure retains original and skips uncertain translation on restore', async () => {
  const { storage, controls } = setup();
  storage.saveAsync.mockImplementationOnce(cb => cb({ status: 'succeeded' }))
    .mockImplementationOnce(cb => cb({ status: 'failed' }));
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('译文已写入');
  expect(rangeBackups()[0].originalOoxml).toBe('<original/>');
  expect(rangeBackups()[0].translatedText).toBeUndefined();
  expect(await restoreOriginalBody()).toEqual({ restored: 0, skipped: 1 });
  expect(controls[0].insertOoxml).not.toHaveBeenCalled();
});

test.each(['Edited', '', 'Original ', 'original'])('text edits during backup persistence cancel replacement: %j', async edited => {
  const { storage, selection } = setup();
  storage.saveAsync.mockImplementationOnce(cb => { selection.text = edited; cb({ status: 'succeeded' }); });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('已更改');
  expect(selection.insertContentControl).not.toHaveBeenCalled();
});

test.each(['selection', 'paragraph', 'body'] as const)('HTML export metadata changes during backup save do not block %s translation', async scope => {
  const ranges = setup();
  ranges[scope].html = '<p id="export-1">Original</p>';
  ranges.storage.saveAsync.mockImplementationOnce(cb => {
    ranges[scope].html = '<p id="export-2"><span>Original</span></p>';
    cb({ status: 'succeeded' });
  });
  await translateDocument(scope, 'ja');
  expect(ranges.controls[0].insertHtml).toHaveBeenCalledWith('<p>译文</p>', 'Replace');
  expect(rangeBackups()[0].originalOoxml).toBe('<original/>');
  expect(await restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
});

test('selection changes during authentication do not redirect translation', async () => {
  const { selection, body, context } = setup();
  jest.mocked(authenticate).mockImplementationOnce(async () => {
    context.document.getSelection.mockReturnValue(body);
    return { token: 'token', user: {} as any };
  });
  await translateDocument('selection', 'ja');
  expect(selection.insertContentControl).toHaveBeenCalled();
  expect(body.insertContentControl).not.toHaveBeenCalled();
});

test.each([null, '', ' '.repeat(3), 'x'.repeat(1000001)])('invalid translation is not written (%#)', async html => {
  const { selection } = setup(); jest.mocked(api).mockResolvedValueOnce({ html });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('译文无效');
  expect(selection.insertContentControl).not.toHaveBeenCalled();
});

test('empty selection, cancellation and network failure never replace content', async () => {
  const { selection } = setup();
  selection.text = '';
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('请先选中');
  selection.text = 'Original';
  await expect(translateDocument('selection', 'ja', () => false)).rejects.toThrow('取消');
  jest.mocked(api).mockRejectedValueOnce(new Error('网络失败'));
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('网络失败');
  expect(selection.insertContentControl).not.toHaveBeenCalled();
});

test('failed settings write preserves previous range backups', async () => {
  const { storage } = setup();
  const original = [{ tag: TAG_PREFIX + 'one', originalOoxml: '<one/>' }];
  await saveRangeBackups(original);
  storage.saveAsync.mockImplementationOnce(cb => cb({ status: 'failed' }));
  await expect(saveRangeBackups([])).rejects.toThrow('保存');
  expect(rangeBackups()).toEqual(original);
});

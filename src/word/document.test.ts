/** @jest-environment jsdom */
import { translateDocument, restoreOriginalBody } from './document';
import { wordPackage } from './test-fixtures/package';
import { api, authenticate } from '../shared/api';
import { rangeBackups, saveRangeBackups, TAG_PREFIX } from './backup';
import { randomUUID } from 'crypto';
jest.mock('../shared/api');

const webPackage = (content = '<w:p><w:r><w:t>Original</w:t></w:r></w:p>') =>
  wordPackage(content).replace('<w:cols w:num="2"/>', '').replace(/<pkg:part pkg:name="\/word\/media\/image1.png">[\s\S]*?<\/pkg:part>/, '');
function setupWeb() {
  const ranges = setup();
  (Office as any).PlatformType = { OfficeOnline: 'OfficeOnline' };
  (Office.context as any).platform = 'OfficeOnline';
  ranges.body.ooxml = webPackage();
  const extra = { text: '', load: jest.fn(), getOoxml: () => ({ value: webPackage('<w:p/>') }) };
  (ranges.context.document as any).sections = { load: jest.fn(), items: [{ getHeader: () => extra, getFooter: () => extra }] };
  return { ...ranges, extra };
}

test.each(['selection', 'paragraph', 'body'] as const)('web %s checks the whole document before requesting translation', async scope => {
  const ranges = setupWeb();
  ranges.body.ooxml = webPackage('<w:tbl/>');
  await expect(translateDocument(scope, 'ja')).rejects.toThrow('表格');
  expect(api).not.toHaveBeenCalled();
  expect(ranges.storage.saveAsync).not.toHaveBeenCalled();
  expect(ranges[scope].insertHtml).not.toHaveBeenCalled();
  expect(ranges[scope].insertOoxml).not.toHaveBeenCalled();
});
test('web allows simple documents and rejects header content', async () => {
  const ranges = setupWeb();
  ranges.extra.text = 'Header';
  await expect(translateDocument('body', 'ja')).rejects.toThrow('页眉');
  ranges.extra.text = '';
  await expect(translateDocument('body', 'ja')).resolves.toEqual({ htmlBackup: false });
});
test('web export failure does not fall back to HTML', async () => {
  const ranges = setupWeb();
  ranges.body.getOoxml.mockImplementation(() => { throw new Error('ooxmlIsMalformed'); });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('兼容性检查');
  expect(api).not.toHaveBeenCalled();
  expect(ranges.selection.insertHtml).not.toHaveBeenCalled();
});
test('web rechecks structures added while the service is running', async () => {
  const ranges = setupWeb();
  jest.mocked(api).mockImplementationOnce(async () => {
    ranges.body.ooxml = webPackage('<w:p><w:r><w:t>Original</w:t><w:drawing/></w:r></w:p>');
    return { html: '<p>译文</p>' } as any;
  });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('图片');
  expect(ranges.selection.insertHtml).not.toHaveBeenCalled();
});
test('web rechecks after backup persistence before replacing text', async () => {
  const ranges = setupWeb();
  ranges.storage.saveAsync.mockImplementationOnce(cb => {
    ranges.body.ooxml = webPackage('<w:tbl/>');
    cb({ status: 'succeeded' });
  });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('表格');
  expect(ranges.selection.insertHtml).not.toHaveBeenCalled();
});

function setup() {
  Object.defineProperty(globalThis, 'crypto', { configurable: true, value: { randomUUID } });
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
      insertHtml: jest.fn((html: string) => { range.html = html; range.text = '译文'; return range; }),
      insertOoxml: jest.fn((xml: string) => { range.ooxml = xml; range.text = '译文'; return range; }),
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
  body.ooxml = wordPackage();
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
  jest.mocked(api).mockImplementation(async (path) => (path === '/api/translate/word'
    ? { paragraphs: ['<p><span id="r0">译文</span></p>'] }
    : { html: '<p>译文</p>' }) as any);
  return { context, selection, paragraph, body, storage, controls };
}

test.each(['selection', 'paragraph', 'body'] as const)('backs up and translates only the %s range', async scope => {
  const ranges = setup();
  ranges[scope].ooxml = scope === 'body' ? wordPackage() : `<${scope}/>`;
  const original = ranges[scope].ooxml;
  await translateDocument(scope, 'ja');
  expect(rangeBackups()[0].originalOoxml).toBe(original);
  expect(rangeBackups()[0].translatedText).toBe('译文');
  if (scope === 'body') {
    expect(ranges.body.insertOoxml).toHaveBeenCalledWith(expect.stringContaining('译文'), 'Replace');
    expect(ranges.body.getHtml).not.toHaveBeenCalled();
    expect(ranges.body.insertHtml).not.toHaveBeenCalled();
  } else expect(ranges[scope].insertHtml).toHaveBeenCalledWith('<p>译文</p>', 'Replace');
  expect(ranges.context.document.body.insertOoxml).not.toHaveBeenCalled();
  expect(ranges.context.trackedObjects.remove).toHaveBeenCalledWith(ranges[scope]);
});

describe.each(['selection', 'paragraph', 'body'] as const)('%s translation with existing document controls', scope => {
  test.each([null, undefined, '', 'template-control'])('ignores unrelated control tag %j', async tag => {
    const ranges = setup();
    const existing = ranges[scope].insertContentControl();
    existing.tag = tag;
    ranges[scope].relation = 'Equal';

    await expect(translateDocument(scope, 'ja')).resolves.toEqual({ htmlBackup: false });

    expect(existing.getRange).not.toHaveBeenCalled();
    expect(existing.delete).not.toHaveBeenCalled();
    expect(scope === 'body' ? ranges.body.insertOoxml : ranges[scope].insertHtml).toHaveBeenCalled();
    expect(rangeBackups()[0].originalOoxml).toBe(scope === 'body' ? wordPackage() : '<original/>');
  });
});

test('untagged controls do not bypass overlap protection for translated content', async () => {
  const { body, controls } = setup();
  const untagged = body.insertContentControl();
  untagged.tag = null;
  await translateDocument('body', 'ja');
  body.relation = 'Contains';

  await expect(translateDocument('body', 'ja')).rejects.toThrow('尚未恢复');

  expect(body.insertOoxml).toHaveBeenCalledTimes(1);
  expect(untagged.delete).not.toHaveBeenCalled();
  expect(controls).toHaveLength(2);
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
  expect(await restoreOriginalBody()).toMatchObject({ restored: 1, skipped: 1 });
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
  expect(await restoreOriginalBody()).toMatchObject({ restored: 0, skipped: 1 });
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
  const incomplete = controls[0];
  expect(await restoreOriginalBody()).toMatchObject({ restored: 0, skipped: 0, cleaned: 1 });
  expect(incomplete.insertOoxml).not.toHaveBeenCalled();
  expect(incomplete.delete).toHaveBeenCalledWith(true);
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
  expect(scope === 'body' ? ranges.body.insertOoxml : ranges[scope].insertHtml).toHaveBeenCalled();
  expect(rangeBackups()[0].originalOoxml).toBe(scope === 'body' ? wordPackage() : '<original/>');
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

test.each(['ooxmlIsMalformed', 'ooxmlIsMalformated'])('OOXML export failure %s falls back to persisted range HTML', async code => {
  const { selection, context, body, controls } = setup();
  let exportPending = false;
  selection.getOoxml.mockImplementation(() => { exportPending = true; return { value: '' }; });
  context.sync.mockImplementation(async () => {
    if (exportPending) { exportPending = false; throw Object.assign(new Error(code), { code }); }
  });
  expect(await translateDocument('selection', 'ja')).toEqual({ htmlBackup: true });
  expect(rangeBackups()[0].originalHtml).toBe('<p>Original</p>');
  expect(rangeBackups()[0].originalOoxml).toBeUndefined();
  body.text = 'Later edits';
  const control = controls[0];
  expect(await restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
  expect(control.insertHtml).toHaveBeenLastCalledWith('<p>Original</p>', 'Replace');
  expect(control.insertOoxml).not.toHaveBeenCalled();
  expect(body.text).toBe('Later edits');
});

test('unrelated export failure does not silently downgrade the backup', async () => {
  const { selection } = setup();
  selection.getOoxml.mockImplementation(() => { throw new Error('AccessDenied'); });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('AccessDenied');
  expect(selection.insertContentControl).not.toHaveBeenCalled();
  expect(rangeBackups()).toEqual([]);
});

test('HTML fallback backup save failure still prevents translation', async () => {
  const { selection, storage } = setup();
  selection.getOoxml.mockImplementation(() => { throw new Error('ooxmlIsMalformed'); });
  storage.saveAsync.mockImplementationOnce(cb => cb({ status: 'failed' }));
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('保存翻译范围备份失败');
  expect(selection.insertContentControl).not.toHaveBeenCalled();
});

test('OOXML error while inserting translation reports the write stage and does not fall back', async () => {
  const { context, selection } = setup();
  let failed = false;
  context.sync.mockImplementation(async () => {
    if (!failed && selection.insertHtml.mock.calls.length) {
      failed = true;
      throw Object.assign(new Error('ooxmlIsMalformed'), { code: 'ooxmlIsMalformed' });
    }
  });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('写入译文');
  expect(rangeBackups()[0].originalOoxml).toBe('<original/>');
  expect(rangeBackups()[0].originalHtml).toBeUndefined();
  expect(rangeBackups()[0].translatedText).toBeUndefined();
});

test('wraps the returned inserted range after import, rather than the original paragraph', async () => {
  const { paragraph, selection } = setup();
  paragraph.insertHtml.mockImplementation(() => { selection.text = '译文'; return selection; });
  paragraph.insertContentControl.mockImplementation(() => { throw new Error('Original boundary unsupported'); });
  await translateDocument('paragraph', 'ja');
  expect(paragraph.insertContentControl).not.toHaveBeenCalled();
  expect(selection.insertContentControl).toHaveBeenCalledTimes(1);
  expect(paragraph.insertHtml.mock.invocationCallOrder[0]).toBeLessThan(selection.insertContentControl.mock.invocationCallOrder[0]);
});

test('property failure removes the new wrapper only and reports the Word API location', async () => {
  const { context, controls, selection } = setup();
  let failedControl: any;
  context.sync.mockImplementation(async () => {
    if (!failedControl && controls[0]?.tag) {
      failedControl = controls[0];
      throw Object.assign(new Error('GeneralException'), {
        code: 'GeneralException', debugInfo: { errorLocation: 'ContentControl.tag' },
      });
    }
  });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('ContentControl.tag');
  expect(failedControl.delete).toHaveBeenCalledWith(true);
  expect(controls).toHaveLength(0);
  expect(selection.text).toBe('译文');
  expect(rangeBackups()[0].originalOoxml).toBe('<original/>');
});

test('retry clears an old incomplete wrapper but leaves its text and backup intact', async () => {
  const { selection, controls, body } = setup();
  const old = selection.insertContentControl();
  old.tag = TAG_PREFIX + 'incomplete';
  await saveRangeBackups([{ tag: old.tag, originalOoxml: '<old/>' }]);
  selection.relation = 'Equal';
  body.text = 'Unrelated saved edits';
  await translateDocument('selection', 'ja');
  expect(old.delete).toHaveBeenCalledWith(true);
  expect(controls).not.toContain(old);
  expect(rangeBackups()[0].originalOoxml).toBe('<old/>');
  expect(body.text).toBe('Unrelated saved edits');
});

test('old recoverable title is corrected when current text already differs from the saved translation', async () => {
  const { selection, controls } = setup();
  await translateDocument('selection', 'ja');
  const control = controls[0];
  control.title = '翻译内容（可恢复原文）';
  selection.text = 'Original restored via Word undo';
  const result = await restoreOriginalBody();
  expect(result).toMatchObject({ restored: 0, skipped: 1 });
  expect(result.details?.[0]).toContain('可能已编辑或撤销');
  expect(control.title).toBe('翻译备份（文字不匹配，未恢复）');
  expect(control.insertOoxml).not.toHaveBeenCalled();
  expect(selection.text).toBe('Original restored via Word undo');
});

test('incomplete legacy wrapper is cleared once without repeated skipped warnings', async () => {
  const { selection } = setup();
  const control = selection.insertContentControl();
  control.tag = TAG_PREFIX + 'old-pending';
  await saveRangeBackups([{ tag: control.tag, originalOoxml: '<original/>' }]);
  expect(await restoreOriginalBody()).toEqual({ restored: 0, skipped: 0, cleaned: 1 });
  expect(selection.text).toBe('Original');
  expect(rangeBackups()).toHaveLength(1);
  expect(await restoreOriginalBody()).toEqual({ restored: 0, skipped: 0 });
});

test('full-body translation restores the exact original OOXML package', async () => {
  const { body, controls } = setup();
  const original = body.ooxml;
  await translateDocument('body', 'zh-Hans');
  expect(api).toHaveBeenCalledWith('/api/translate/word', 'token', {
    paragraphs: ['<p><span id="r0">Original</span></p>'], to: 'zh-Hans',
  });
  const control = controls[0];
  expect(await restoreOriginalBody()).toEqual({ restored: 1, skipped: 0 });
  expect(control.insertOoxml).toHaveBeenCalledWith(original, 'Replace');
  expect(body.ooxml).toBe(original);
  expect(body.getHtml).not.toHaveBeenCalled();
  expect(body.insertHtml).not.toHaveBeenCalled();
});

test.each(['export', 'malformed', 'response', 'backup', 'cancel', 'edit', 'insert'])('full-body %s failure never falls back to HTML', async failure => {
  const { body, storage } = setup();
  if (failure === 'export') body.getOoxml.mockImplementation(() => { throw new Error('ooxmlIsMalformed'); });
  if (failure === 'malformed') body.ooxml = '<broken>';
  if (failure === 'response') jest.mocked(api).mockResolvedValueOnce({ paragraphs: [] });
  if (failure === 'backup') storage.saveAsync.mockImplementationOnce(cb => cb({ status: 'failed' }));
  if (failure === 'edit') storage.saveAsync.mockImplementationOnce(cb => { body.text = 'User edits'; cb({ status: 'succeeded' }); });
  if (failure === 'insert') body.insertOoxml.mockImplementation(() => { throw new Error('ooxmlIsMalformed'); });
  await expect(translateDocument('body', 'ja', () => failure !== 'cancel')).rejects.toThrow();
  expect(body.getHtml).not.toHaveBeenCalled();
  expect(body.insertHtml).not.toHaveBeenCalled();
  if (failure !== 'insert') expect(body.insertOoxml).not.toHaveBeenCalled();
  else expect(rangeBackups()[0].originalOoxml).toBe(wordPackage());
});

test('layout-only edits during backup persistence are retained in the inserted full body', async () => {
  const { body, storage } = setup();
  storage.saveAsync.mockImplementationOnce(cb => {
    body.ooxml = body.ooxml.replace('11906', '15000');
    cb({ status: 'succeeded' });
  });
  await expect(translateDocument('body', 'ja')).resolves.toEqual({ htmlBackup: false });
  expect(body.insertOoxml).toHaveBeenCalledWith(expect.stringContaining('15000'), 'Replace');
  expect(body.insertHtml).not.toHaveBeenCalled();
});

test('changing export metadata on every read does not falsely report user edits', async () => {
  const { body } = setup();
  let exportId = 0;
  body.getOoxml.mockImplementation(() => ({ value: wordPackage()
    .replace('<w:p>', `<w:p xmlns:w14="urn:word14" w14:paraId="${++exportId}" w14:textId="${exportId}">`)
    .replace('pkg:name="/word/document.xml"', `pkg:name="/word/document.xml" pkg:padding="${exportId}"`) }));
  await expect(translateDocument('body', 'ja')).resolves.toEqual({ htmlBackup: false });
  expect(body.insertOoxml).toHaveBeenCalledWith(expect.stringContaining(`w14:paraId="${exportId}"`), 'Replace');
  expect(body.insertHtml).not.toHaveBeenCalled();
});

test('backs up and translates the latest layout after the network request', async () => {
  const { body, controls } = setup();
  let latest = '';
  jest.mocked(api).mockImplementationOnce(async () => {
    latest = body.ooxml.replace('11906', '15000').replace('aW1hZ2U=', 'bmV3');
    body.ooxml = latest;
    return { paragraphs: ['<p><span id="r0">译文</span></p>'] } as any;
  });
  await translateDocument('body', 'ja');
  expect(rangeBackups()[0].originalOoxml).toBe(latest);
  expect(body.insertOoxml).toHaveBeenCalledWith(expect.stringContaining('bmV3'), 'Replace');
  const control = controls[0];
  await restoreOriginalBody();
  expect(control.insertOoxml).toHaveBeenCalledWith(latest, 'Replace');
});

test('text-box edits absent from Range.text still prevent full-body replacement', async () => {
  const { body, storage } = setup();
  body.ooxml = wordPackage('<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>');
  storage.saveAsync.mockImplementationOnce(cb => {
    body.ooxml = body.ooxml.replace('>Box<', '>Edited<');
    cb({ status: 'succeeded' });
  });
  await expect(translateDocument('body', 'ja')).rejects.toThrow('正文文字已更改');
  expect(body.insertOoxml).not.toHaveBeenCalled();
  expect(body.insertHtml).not.toHaveBeenCalled();
});

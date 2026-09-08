import { translateDocument, restoreOriginalBody } from './document';
import { api, authenticate } from '../shared/api';
jest.mock('../shared/api');

function setup() {
  const values = new Map<string, string>();
  const storage = { get: jest.fn((key: string) => values.get(key)),
    set: jest.fn((key: string, value: string) => values.set(key, value)),
    remove: jest.fn((key: string) => values.delete(key)),
    saveAsync: jest.fn((cb: Function) => cb({ status: 'succeeded' })) };
  (globalThis as any).Office = { context: { document: { settings: storage } }, AsyncResultStatus: { Succeeded: 'succeeded' } };
  const makeRange = () => ({ text: 'Original', load: jest.fn(), getHtml: jest.fn(() => ({ value: '<p>Original</p>' })),
    getOoxml: jest.fn(() => ({ value: '<original/>' })), insertHtml: jest.fn() });
  const paragraph = makeRange();
  const body = makeRange();
  const selection = { ...makeRange(), paragraphs: { getFirst: () => ({ getRange: () => paragraph }) } };
  const context = { document: { body: { getRange: () => body, insertOoxml: jest.fn() }, getSelection: jest.fn(() => selection) },
    sync: jest.fn(async () => {}), trackedObjects: { add: jest.fn(), remove: jest.fn() } };
  (globalThis as any).Word = { run: (fn: Function) => fn(context), InsertLocation: { replace: 'Replace' } };
  jest.mocked(authenticate).mockResolvedValue({ token: 'token', user: {} as any });
  jest.mocked(api).mockResolvedValue({ html: '<p>译文</p>' });
  return { context, selection, paragraph, body, storage };
}

test.each(['selection', 'paragraph', 'body'] as const)('replaces only the captured %s range using the shared authenticated API', async scope => {
  const ranges = setup();
  await translateDocument(scope, 'ja');
  expect(api).toHaveBeenCalledWith('/api/translate', 'token', { html: '<p>Original</p>', to: 'ja' });
  for (const name of ['selection', 'paragraph', 'body'] as const) {
    expect(ranges[name].insertHtml).toHaveBeenCalledTimes(name === scope ? 1 : 0);
  }
  expect(ranges.context.trackedObjects.remove).toHaveBeenCalledWith(ranges[scope]);
});

test('rejects an empty selection without authentication or a write', async () => {
  const { selection } = setup(); selection.text = '  ';
  await expect(translateDocument('selection', 'zh-Hans')).rejects.toThrow('请先选中');
  expect(authenticate).not.toHaveBeenCalled();
  expect(selection.insertHtml).not.toHaveBeenCalled();
});

test.each(['selection', 'paragraph', 'body'] as const)('export serialization changes do not block unchanged %s text', async scope => {
  const ranges = setup();
  const range = ranges[scope];
  range.getHtml.mockReturnValueOnce({ value: '<p id="export-1">Original</p>' });
  jest.mocked(api).mockImplementationOnce(async () => {
    range.getHtml.mockReturnValue({ value: '<p id="export-2"><span>Original</span></p>' });
    return { html: '<p>译文</p>' } as any;
  });
  await translateDocument(scope, 'zh-Hans');
  expect(range.insertHtml).toHaveBeenCalledWith('<p>译文</p>', 'Replace');
  expect(range.getHtml).toHaveBeenCalledTimes(1);
});

test.each(['Edited', '', 'Original ', 'original'])('reloads text and rejects edits even when HTML is unchanged: %j', async edited => {
  const { selection, context } = setup();
  context.sync.mockImplementation(async () => {
    if (selection.load.mock.calls.length === 2) selection.text = edited;
  });
  await expect(translateDocument('selection', 'zh-Hans')).rejects.toThrow('原文已更改');
  expect(selection.load).toHaveBeenCalledTimes(2);
  expect(selection.insertHtml).not.toHaveBeenCalled();
});

test.each(['selection', 'paragraph', 'body'] as const)('OOXML metadata changes when saving backup do not block %s translation', async scope => {
  const ranges = setup();
  ranges.storage.saveAsync.mockImplementationOnce(cb => {
    ranges[scope].getOoxml.mockReturnValue({ value: '<package-with-updated-settings/>' });
    cb({ status: 'succeeded' });
  });
  await translateDocument(scope, 'ja');
  expect(ranges[scope].insertHtml).toHaveBeenCalledWith('<p>译文</p>', 'Replace');
  expect(ranges.storage.get('wordTranslation.originalBody.v1')).toBe('<original/>');
});

test('a selection change during authentication does not redirect the write', async () => {
  const { selection, body, context } = setup();
  jest.mocked(authenticate).mockImplementationOnce(async () => {
    context.document.getSelection.mockReturnValue(body as any);
    return { token: 'token', user: {} as any };
  });
  await translateDocument('selection', 'zh-Hans');
  expect(selection.insertHtml).toHaveBeenCalled();
  expect(body.insertHtml).not.toHaveBeenCalled();
});

test('cancellation prevents a pending translation from writing', async () => {
  const { body } = setup();
  await expect(translateDocument('body', 'zh-Hans', () => false)).rejects.toThrow('取消翻译');
  expect(body.insertHtml).not.toHaveBeenCalled();
});

test.each([null, '', ' '.repeat(3), 'x'.repeat(1000001)])('invalid translation leaves the document unchanged', async html => {
  const { selection } = setup(); jest.mocked(api).mockResolvedValueOnce({ html });
  await expect(translateDocument('selection', 'zh-Hans')).rejects.toThrow('译文无效');
  expect(selection.insertHtml).not.toHaveBeenCalled();
});

test('network failure releases the tracked range without writing', async () => {
  const { selection, context } = setup(); jest.mocked(api).mockRejectedValueOnce(new Error('网络失败'));
  await expect(translateDocument('selection', 'zh-Hans')).rejects.toThrow('网络失败');
  expect(selection.insertHtml).not.toHaveBeenCalled();
  expect(context.trackedObjects.remove).toHaveBeenCalledWith(selection);
});

test('saves the entire original body before replacing a selection and preserves the first backup', async () => {
  const { body, storage, selection } = setup();
  body.getOoxml.mockReturnValue({ value: '<whole-original-body/>' });
  storage.saveAsync.mockImplementation(cb => {
    expect(selection.insertHtml).not.toHaveBeenCalled();
    cb({ status: 'succeeded' });
  });
  await translateDocument('selection', 'ja');
  body.getOoxml.mockReturnValue({ value: '<translated-body/>' });
  await translateDocument('selection', 'ja');
  expect(storage.set).toHaveBeenCalledTimes(1);
  expect(storage.set).toHaveBeenCalledWith('wordTranslation.originalBody.v1', '<whole-original-body/>');
});

test('backup save failure prevents replacement and allows retry', async () => {
  const { storage, selection } = setup();
  storage.saveAsync.mockImplementationOnce(cb => cb({ status: 'failed' }));
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('保存原文备份失败');
  expect(selection.insertHtml).not.toHaveBeenCalled();
  expect(storage.get('wordTranslation.originalBody.v1')).toBeUndefined();
  await translateDocument('selection', 'ja');
  expect(selection.insertHtml).toHaveBeenCalledTimes(1);
});

test('edits during backup persistence prevent translation from overwriting the selection', async () => {
  const { storage, selection } = setup();
  storage.saveAsync.mockImplementationOnce(cb => {
    selection.text = 'Edited';
    cb({ status: 'succeeded' });
  });
  await expect(translateDocument('selection', 'ja')).rejects.toThrow('原文已更改');
  expect(selection.insertHtml).not.toHaveBeenCalled();
});

test('restores persisted OOXML without authentication and retains backup after restoration', async () => {
  const { storage, context } = setup();
  storage.set('wordTranslation.originalBody.v1', '<persisted-original/>');
  await restoreOriginalBody();
  expect(context.document.body.insertOoxml).toHaveBeenCalledWith('<persisted-original/>', 'Replace');
  expect(authenticate).not.toHaveBeenCalled();
  expect(api).not.toHaveBeenCalled();
  expect(storage.get('wordTranslation.originalBody.v1')).toBe('<persisted-original/>');
});

test('missing backup leaves the document unchanged', async () => {
  const { context } = setup();
  await expect(restoreOriginalBody()).rejects.toThrow('没有保存的原文备份');
  expect(context.document.body.insertOoxml).not.toHaveBeenCalled();
});

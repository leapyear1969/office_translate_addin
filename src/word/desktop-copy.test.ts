/** @jest-environment jsdom */
import { randomUUID } from 'crypto';
import { COPY_KEY, exportDocument, prepareDesktopCopy } from './desktop-copy';

function setup() {
  Object.defineProperty(globalThis, 'crypto', { configurable: true, value: { randomUUID } });
  const sourceRange = { text: 'Original', load: jest.fn(), insertBookmark: jest.fn() };
  const copyRange = { text: 'Original', load: jest.fn() };
  const settings = { get: jest.fn(), add: jest.fn() };
  const source = { body: { getRange: () => sourceRange }, getSelection: () => ({ ...sourceRange,
    paragraphs: { getFirst: () => ({ getRange: () => sourceRange }) } }), deleteBookmark: jest.fn() };
  const copy = { body: { getRange: () => copyRange }, getBookmarkRange: jest.fn(() => copyRange),
    deleteBookmark: jest.fn(), settings, open: jest.fn() };
  const file = { sliceCount: 2, getSliceAsync: jest.fn((i, cb) => cb({ status: 'ok', value: { data: i ? [3, 4, 5] : [1, 2] } })),
    closeAsync: jest.fn(cb => cb({ status: 'ok' })) };
  const getFileAsync = jest.fn((_type, _options, cb) => cb({ status: 'ok', value: file }));
  (globalThis as any).Office = { AsyncResultStatus: { Succeeded: 'ok' }, FileType: { Compressed: 'compressed' },
    context: { document: { settings, getFileAsync }, requirements: { isSetSupported: jest.fn(() => true) } } };
  const context: any = { document: source, application: { createDocument: jest.fn(() => copy) },
    sync: jest.fn(async () => {}), trackedObjects: { add: jest.fn(), remove: jest.fn() } };
  return { context, source, sourceRange, copy, copyRange, settings, file, getFileAsync };
}

test('full DOCX export joins non-aligned slices and closes the file', async () => {
  const { file } = setup();
  expect(await exportDocument()).toBe('AQIDBAU=');
  expect(file.closeAsync).toHaveBeenCalledTimes(1);
});

test.each(['selection', 'paragraph', 'body'] as const)('creates full copy and binds %s to that copy', async scope => {
  const { context, source, copy, copyRange, sourceRange } = setup();
  const prepared = await prepareDesktopCopy(context, scope);
  expect(context.application.createDocument).toHaveBeenCalledWith('AQIDBAU=');
  expect(prepared.document).toBe(copy);
  expect(prepared.range).toBe(copyRange);
  expect(copy.settings.add).toHaveBeenCalledWith(COPY_KEY, true);
  if (scope !== 'body') {
    const bookmark = sourceRange.insertBookmark.mock.calls[0][0];
    expect(copy.getBookmarkRange).toHaveBeenCalledWith(bookmark);
    expect(source.deleteBookmark).toHaveBeenCalledWith(bookmark);
    expect(copy.deleteBookmark).toHaveBeenCalledWith(bookmark);
  } else expect(sourceRange.insertBookmark).not.toHaveBeenCalled();
  await prepared.finish();
  expect(copy.open).toHaveBeenCalledTimes(1);
});

test('existing translation copy is reused', async () => {
  const { context, settings, source, getFileAsync } = setup();
  settings.get.mockReturnValue(true);
  const prepared = await prepareDesktopCopy(context, 'body');
  expect(prepared.document).toBe(source);
  expect(getFileAsync).not.toHaveBeenCalled();
  expect(context.application.createDocument).not.toHaveBeenCalled();
});

test('failed export cleans source bookmark and closes the file without creating a copy', async () => {
  const { context, file, source } = setup();
  file.getSliceAsync.mockImplementationOnce((_index, cb) => cb({ status: 'failed', error: { message: 'export failed' } }));
  await expect(prepareDesktopCopy(context, 'selection')).rejects.toThrow('export failed');
  expect(source.deleteBookmark).toHaveBeenCalledTimes(1);
  expect(file.closeAsync).toHaveBeenCalledTimes(1);
  expect(context.application.createDocument).not.toHaveBeenCalled();
});

test('unsupported host stops before exporting or adding bookmarks', async () => {
  const { context, getFileAsync, sourceRange } = setup();
  jest.mocked(Office.context.requirements.isSetSupported).mockReturnValue(false);
  await expect(prepareDesktopCopy(context, 'selection')).rejects.toThrow('更新桌面 Word');
  expect(getFileAsync).not.toHaveBeenCalled();
  expect(sourceRange.insertBookmark).not.toHaveBeenCalled();
});

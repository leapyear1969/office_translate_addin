import { translateDocument } from './document';
import { api, authenticate } from '../shared/api';
jest.mock('../shared/api');

function setup() {
  const makeRange = () => ({ text: 'Original', load: jest.fn(), getHtml: jest.fn(() => ({ value: '<p>Original</p>' })),
    getOoxml: jest.fn(() => ({ value: '<original/>' })), insertHtml: jest.fn() });
  const paragraph = makeRange();
  const body = makeRange();
  const selection = { ...makeRange(), paragraphs: { getFirst: () => ({ getRange: () => paragraph }) } };
  const context = { document: { body: { getRange: () => body }, getSelection: jest.fn(() => selection) },
    sync: jest.fn(async () => {}), trackedObjects: { add: jest.fn(), remove: jest.fn() } };
  (globalThis as any).Word = { run: (fn: Function) => fn(context), InsertLocation: { replace: 'Replace' } };
  jest.mocked(authenticate).mockResolvedValue({ token: 'token', user: {} as any });
  jest.mocked(api).mockResolvedValue({ html: '<p>译文</p>' });
  return { context, selection, paragraph, body };
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

test('does not overwrite text or formatting edited while translation is pending', async () => {
  const { selection } = setup();
  selection.getOoxml.mockReturnValueOnce({ value: '<original/>' }).mockReturnValueOnce({ value: '<edited/>' });
  await expect(translateDocument('selection', 'zh-Hans')).rejects.toThrow('原文已更改');
  expect(selection.insertHtml).not.toHaveBeenCalled();
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

test('rejects a stale automatic offer before sending content', async () => {
  const { body } = setup();
  await expect(translateDocument('body', 'zh-Hans', '<p>Older</p>')).rejects.toThrow('文档已更改');
  expect(api).not.toHaveBeenCalled();
  expect(body.insertHtml).not.toHaveBeenCalled();
});

test('closing the pane cancels an automatic write', async () => {
  const { body } = setup();
  await expect(translateDocument('body', 'zh-Hans', '<p>Original</p>', () => false)).rejects.toThrow('取消自动翻译');
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

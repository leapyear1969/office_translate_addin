/** @jest-environment jsdom */
import { capturePreview, translatePreview } from './preview';
import { api, authenticate } from '../shared/api';
jest.mock('../shared/api');

function setup() {
  const range = { text: 'Original', load: jest.fn(), insertText: jest.fn() };
  const selection = { ...range, paragraphs: { getFirst: () => ({ getRange: () => range }) } };
  const context = { document: { getSelection: jest.fn(() => selection) }, sync: jest.fn(async () => {}),
    trackedObjects: { add: jest.fn(), remove: jest.fn() } };
  (globalThis as any).Word = { InsertLocation: { replace: 'Replace' }, run: jest.fn(async (...args) => args[args.length - 1](context)) };
  return { range, selection, context };
}

test('captures a paragraph without writing and inserts into the captured range after focus moves', async () => {
  const { range, context } = setup();
  const preview = await capturePreview('paragraph');
  expect(preview.text).toBe('Original');
  expect(range.insertText).not.toHaveBeenCalled();
  context.document.getSelection.mockImplementation(() => { throw new Error('selection moved'); });
  await preview.insert('Translated', () => true);
  expect(range.insertText).toHaveBeenCalledWith('Translated', 'Replace');
  await preview.release(); await preview.release();
  expect(context.trackedObjects.remove).toHaveBeenCalledTimes(1);
});

test('changed source or cancelled session never overwrites the document', async () => {
  const { selection } = setup();
  const preview = await capturePreview('selection');
  selection.text = 'Edited';
  await expect(preview.insert('Translation', () => true)).rejects.toThrow('原文范围已更改');
  selection.text = 'Original';
  await expect(preview.insert('Translation', () => false)).rejects.toThrow('已取消');
  expect(selection.insertText).not.toHaveBeenCalled();
  await preview.release();
  await expect(preview.insert('Translation', () => true)).rejects.toThrow('重新选择');
});

test('empty selections are rejected without retaining a Word range', async () => {
  const { selection, context } = setup();
  selection.text = '';
  await expect(capturePreview('selection')).rejects.toThrow('请先选中');
  expect(context.trackedObjects.add).not.toHaveBeenCalled();
});

test('preview preserves literal markup and newlines as text', async () => {
  jest.mocked(authenticate).mockResolvedValue({ token: 'token' } as any);
  jest.mocked(api).mockResolvedValue({ html: '<div>&lt;b&gt;译文&lt;/b&gt;\n第二行 &amp; 第三行</div>' });
  expect(await translatePreview('<b>source</b>\nnext & last', 'en')).toBe('<b>译文</b>\n第二行 & 第三行');
  expect(api).toHaveBeenLastCalledWith('/api/translate', 'token', { html: '<div>&lt;b&gt;source&lt;/b&gt;\nnext &amp; last</div>', to: 'en' });
});

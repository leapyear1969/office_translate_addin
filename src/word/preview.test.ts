/** @jest-environment jsdom */
import { capturePreview, translatePreview, EmptySelectionError } from './preview';
import { api, authenticate } from '../shared/api';
import { wordPackage } from './test-fixtures/package';
jest.mock('../shared/api');

function setup() {
  const control = { id: 1, isNullObject: true, load: jest.fn(), getOoxml: jest.fn(() => ({ value: wordPackage('<w:p/>') })), insertText: jest.fn() };
  const contained = { items: [] as typeof control[], load: jest.fn() };
  const range = { text: 'Original', isEmpty: false, load: jest.fn(), insertText: jest.fn(), getOoxml: jest.fn(() => ({ value: wordPackage('<w:p/>') })), parentContentControlOrNullObject: control, contentControls: contained };
  const paragraph = { getRange: jest.fn((location?: string) => location === 'Content' ? range : {
    ...range, text: 'Original\r', insertText: jest.fn(() => { throw new Error('GeneralException'); }),
  }) };
  const selection = { ...range, paragraphs: { getFirst: () => paragraph } };
  const context = { document: { getSelection: jest.fn(() => selection) }, sync: jest.fn(async () => {}),
    trackedObjects: { add: jest.fn(), remove: jest.fn() } };
  (globalThis as any).Word = { InsertLocation: { replace: 'Replace' }, run: jest.fn(async (...args) => args[args.length - 1](context)) };
  return { range, selection, context, paragraph, control, contained };
}

test('captures a paragraph without writing and inserts into the captured range after focus moves', async () => {
  const { range, context, paragraph } = setup();
  const preview = await capturePreview('paragraph');
  expect(paragraph.getRange).toHaveBeenCalledWith('Content');
  expect(preview.text).toBe('Original');
  expect(range.insertText).not.toHaveBeenCalled();
  context.document.getSelection.mockImplementation(() => { throw new Error('selection moved'); });
  await preview.insert('Translated', () => true);
  expect(range.insertText).toHaveBeenCalledWith('Translated', 'Replace');
  await preview.release(); await preview.release();
  expect(context.trackedObjects.remove).toHaveBeenCalledTimes(1);
});

test('replacement errors identify the failing step and host location without retrying the write', async () => {
  const { range, context } = setup();
  const preview = await capturePreview('paragraph');
  context.sync.mockResolvedValueOnce(undefined).mockRejectedValueOnce(Object.assign(new Error('GeneralException'), {
    code: 'GeneralException', debugInfo: { errorLocation: 'Range.insertText' },
  }));
  await expect(preview.insert('译文', () => true)).rejects.toThrow('替换原文失败：GeneralException（Range.insertText）');
  expect(range.insertText).toHaveBeenCalledTimes(1);
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

test('a collapsed cursor returns guidance without requesting OOXML that can throw GeneralException', async () => {
  const { selection, context } = setup();
  selection.text = '';
  selection.isEmpty = true;
  selection.getOoxml.mockImplementation(() => { throw new Error('GeneralException'); });
  await expect(capturePreview('selection')).rejects.toBeInstanceOf(EmptySelectionError);
  expect(selection.getOoxml).not.toHaveBeenCalled();
  expect(context.trackedObjects.add).not.toHaveBeenCalled();
});

test('preview preserves literal markup and newlines as text', async () => {
  jest.mocked(authenticate).mockResolvedValue({ token: 'token' } as any);
  jest.mocked(api).mockResolvedValue({ html: '<div>&lt;b&gt;译文&lt;/b&gt;\n第二行 &amp; 第三行</div>' });
  expect(await translatePreview('<b>source</b>\nnext & last', 'en')).toBe('<b>译文</b>\n第二行 & 第三行');
  expect(api).toHaveBeenLastCalledWith('/api/translate', 'token', { html: '<div>&lt;b&gt;source&lt;/b&gt;\nnext &amp; last</div>', to: 'en' });
});

test('preview reads empty-text template controls and checks hidden source edits before insertion', async () => {
  const { selection, control, contained } = setup();
  selection.text = '';
  const xml = wordPackage('<w:p><w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:r><w:t>Sample</w:t></w:r></w:sdtContent></w:sdt></w:p>');
  selection.getOoxml.mockReturnValue({ value: xml });
  contained.items = [control];
  control.getOoxml.mockImplementation(() => selection.getOoxml());
  const preview = await capturePreview('selection');
  expect(preview.text).toBe('Sample');
  selection.getOoxml.mockReturnValue({ value: xml.replace('Sample', 'Changed') });
  await expect(preview.insert('译文', () => true)).rejects.toThrow('原文范围已更改');
  expect(selection.insertText).not.toHaveBeenCalled();
  selection.getOoxml.mockReturnValue({ value: xml });
  await preview.insert('译文', () => true);
  expect(control.insertText).toHaveBeenCalledWith('译文', 'Replace');
});

// Modern Living uses hidden rich-text placeholders containing three paragraphs.
const templatePlaceholder = wordPackage('<w:sdt><w:sdtPr><w:id w:val="1620411488"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
  ['First paragraph.', 'Second paragraph.', 'Third paragraph.'].map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join('') +
  '</w:sdtContent></w:sdt>');

test.each(['parent', 'contained'] as const)('visible placeholder text uses its %s control, even when Range.text is populated', async location => {
  const { range, control, contained } = setup();
  range.text = 'First paragraph.\rSecond paragraph.\rThird paragraph.';
  range.getOoxml.mockReturnValue({ value: templatePlaceholder });
  control.getOoxml.mockReturnValue({ value: templatePlaceholder });
  if (location === 'parent') control.isNullObject = false;
  else contained.items = [control];
  range.insertText.mockImplementation(() => { throw new Error('GeneralException'); });
  const preview = await capturePreview('paragraph');
  expect(preview.text).toBe(range.text);
  await preview.insert('译文', () => true);
  expect(control.insertText).toHaveBeenCalledWith('译文', 'Replace');
  expect(range.insertText).not.toHaveBeenCalled();
});

test('unresolved template prompts stop before writing instead of falling back to Range.insertText', async () => {
  const { range } = setup();
  range.text = '';
  range.getOoxml.mockReturnValue({ value: templatePlaceholder });
  await expect(capturePreview('paragraph')).rejects.toThrow('无法定位模板占位内容控件');
  expect(range.insertText).not.toHaveBeenCalled();
});

test('ordinary content controls retain partial-range replacement', async () => {
  const { range, control } = setup();
  control.isNullObject = false;
  control.getOoxml.mockReturnValue({ value: templatePlaceholder.replace('<w:showingPlcHdr/>', '') });
  const preview = await capturePreview('paragraph');
  await preview.insert('译文', () => true);
  expect(range.insertText).toHaveBeenCalledWith('译文', 'Replace');
  expect(control.insertText).not.toHaveBeenCalled();
});

test('multi-paragraph template prompts write through the captured content control and detect hidden edits', async () => {
  const { range, control, context } = setup();
  range.text = '';
  range.getOoxml.mockReturnValue({ value: templatePlaceholder });
  control.isNullObject = false;
  control.getOoxml.mockReturnValue({ value: templatePlaceholder });
  range.insertText.mockImplementation(() => { throw new Error('GeneralException'); });
  const preview = await capturePreview('paragraph');
  expect(preview.text).toBe('First paragraph.\nSecond paragraph.\nThird paragraph.');
  control.getOoxml.mockReturnValue({ value: templatePlaceholder.replace('Third', 'Edited') });
  await expect(preview.insert('译文', () => true)).rejects.toThrow('原文范围已更改');
  expect(control.insertText).not.toHaveBeenCalled();
  control.getOoxml.mockReturnValue({ value: templatePlaceholder });
  await preview.insert('第一段\n第二段\n第三段', () => true);
  expect(control.insertText).toHaveBeenCalledWith('第一段\n第二段\n第三段', 'Replace');
  expect(range.insertText).not.toHaveBeenCalled();
  await preview.release();
  expect(context.trackedObjects.remove).toHaveBeenCalledWith(control);
});

test('a partial placeholder preview cannot overwrite the larger enclosing control', async () => {
  const { selection, control, context } = setup();
  selection.text = '';
  selection.getOoxml.mockReturnValue({ value: wordPackage('<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>') });
  control.isNullObject = false;
  control.getOoxml.mockReturnValue({ value: templatePlaceholder });
  await expect(capturePreview('selection')).rejects.toThrow('仅包含模板占位内容的一部分');
  expect(control.insertText).not.toHaveBeenCalled();
  expect(context.trackedObjects.add).not.toHaveBeenCalled();
});

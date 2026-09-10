/** @jest-environment jsdom */
import { copyTranslation } from './clipboard';

const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
const originalExecCommand = Object.getOwnPropertyDescriptor(document, 'execCommand');
let field: HTMLTextAreaElement;
let button: HTMLButtonElement;
let writeText: jest.Mock;
let execCommand: jest.Mock;

beforeEach(() => {
  document.body.innerHTML = '<textarea readonly></textarea><button>Copy</button>';
  field = document.querySelector('textarea')!;
  button = document.querySelector('button')!;
  field.value = '你好，Jason！\n第二行 🌍';
  field.setSelectionRange(1, 3, 'backward');
  button.focus();
  writeText = jest.fn().mockResolvedValue(undefined);
  execCommand = jest.fn().mockReturnValue(false);
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
  Object.defineProperty(document, 'execCommand', { configurable: true, value: execCommand });
});

afterEach(() => {
  if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
  else Reflect.deleteProperty(navigator, 'clipboard');
  if (originalExecCommand) Object.defineProperty(document, 'execCommand', originalExecCommand);
  else Reflect.deleteProperty(document, 'execCommand');
});

test('copies the full readonly translation synchronously even when iframe Clipboard API is denied', async () => {
  writeText.mockRejectedValue(new DOMException('Permissions policy', 'NotAllowedError'));
  execCommand.mockImplementation(command => {
    expect(command).toBe('copy');
    expect(document.activeElement).toBe(field);
    expect(field.value.slice(field.selectionStart, field.selectionEnd)).toBe(field.value);
    return true;
  });
  const result = copyTranslation(field);
  expect(execCommand).toHaveBeenCalledTimes(1);
  expect(await result).toBe(true);
  expect(writeText).not.toHaveBeenCalled();
  expect(document.activeElement).toBe(button);
  expect([field.selectionStart, field.selectionEnd, field.selectionDirection]).toEqual([1, 3, 'backward']);
});

test.each(['false', 'throws', 'missing'])('uses Clipboard API when copy command is %s', async mode => {
  if (mode === 'throws') execCommand.mockImplementation(() => { throw new Error('Unsupported'); });
  if (mode === 'missing') Reflect.deleteProperty(document, 'execCommand');
  expect(await copyTranslation(field)).toBe(true);
  expect(writeText).toHaveBeenCalledWith(field.value);
  expect(document.activeElement).toBe(button);
});

test.each(['denied', 'missing'])('selects translation for manual copy when both paths fail and Clipboard API is %s', async mode => {
  if (mode === 'missing') Reflect.deleteProperty(navigator, 'clipboard');
  else writeText.mockRejectedValue(new DOMException('Denied', 'NotAllowedError'));
  expect(await copyTranslation(field)).toBe(false);
  expect(document.activeElement).toBe(field);
  expect(field.selectionStart).toBe(0);
  expect(field.selectionEnd).toBe(field.value.length);
});

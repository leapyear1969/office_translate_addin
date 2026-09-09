/** @jest-environment jsdom */
import { rememberPaneWidth } from './pane-width';

let cleanup: (() => void) | undefined;
let setWidth: jest.Mock;
beforeEach(() => {
  jest.useFakeTimers();
  localStorage.clear();
  setWidth = jest.fn();
  (globalThis as any).Office = {
    context: { platform: 'PC' }, PlatformType: { OfficeOnline: 'web' },
    extensionLifeCycle: { taskpane: { setWidth } },
  };
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: 400 });
});
afterEach(() => { cleanup?.(); localStorage.clear(); jest.useRealTimers(); });
function drag(width: number) {
  Object.defineProperty(window, 'innerWidth', { configurable: true, value: width });
  window.dispatchEvent(new Event('resize'));
  jest.advanceTimersByTime(250);
}
test('restores the width after closing and reopening', () => {
  cleanup = rememberPaneWidth();
  drag(450);
  cleanup();
  cleanup = rememberPaneWidth();
  expect(setWidth).toHaveBeenLastCalledWith(450);
});
test('corrects narrow panes without replacing the saved width', () => {
  cleanup = rememberPaneWidth();
  drag(420);
  drag(200);
  expect(setWidth).toHaveBeenLastCalledWith(330);
  expect(localStorage.getItem('word-translator.pane-width')).toBe('420');
});
test('clamps restored widths to the web host limit', () => {
  Office.context.platform = Office.PlatformType.OfficeOnline;
  localStorage.setItem('word-translator.pane-width', '700');
  cleanup = rememberPaneWidth();
  expect(setWidth).toHaveBeenCalledWith(500);
});
test('remains usable when the host cannot resize the pane', () => {
  delete (Office as any).extensionLifeCycle;
  localStorage.setItem('word-translator.pane-width', '420');
  expect(() => { cleanup = rememberPaneWidth(); drag(430); }).not.toThrow();
});

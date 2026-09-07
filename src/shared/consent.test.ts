/**
 * @jest-environment jsdom
 * @jest-environment-options {"url":"https://addin.example"}
 */
jest.mock('./api', () => ({ api: jest.fn() }));
import { api } from './api';
import { requestConsent } from './consent';
const url = 'https://addin.example/consent.html#state=test';
beforeEach(() => { (api as jest.Mock).mockResolvedValue({ url }); });
test('opens the external browser and returns a fallback link without waiting for completion', async () => {
  const openBrowserWindow = jest.fn();
  (global as any).Office = { context: { ui: { openBrowserWindow } } };
  await expect(requestConsent('token')).resolves.toBe(url);
  expect(openBrowserWindow).toHaveBeenCalledWith(url);
});
test('unsupported hosts can still offer a clickable link', async () => {
  (global as any).Office = { context: { ui: {} } };
  await expect(requestConsent('token')).resolves.toBe(url);
});
test('does not open a foreign origin', async () => {
  const openBrowserWindow = jest.fn();
  (global as any).Office = { context: { ui: { openBrowserWindow } } };
  (api as jest.Mock).mockResolvedValue({ url: 'https://other.example/consent.html' });
  await expect(requestConsent('token')).rejects.toThrow('地址无效');
  expect(openBrowserWindow).not.toHaveBeenCalled();
});

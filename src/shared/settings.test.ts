import { normalizeSettings, shouldOfferTranslation, saveSettings, loadSettings } from './settings';
test('uses safe defaults and normalizes invalid stored settings', () => {
  expect(normalizeSettings(null)).toEqual({ mode: 'ask', target: 'zh-Hans', excluded: [] });
  expect(normalizeSettings({ mode: 'bogus', target: '../../x', excluded: ['en', 'en', 'bad'] })).toEqual({ mode: 'ask', target: 'zh-Hans', excluded: ['en'] });
});
test('respects never mode, target language, ignored languages and uncertain detection', () => {
  const settings = normalizeSettings({ mode: 'always', target: 'zh-Hans', excluded: ['en'] });
  expect(shouldOfferTranslation(settings, 'en', 1)).toBe(false);
  expect(shouldOfferTranslation(settings, 'zh-Hans', 1)).toBe(false);
  expect(shouldOfferTranslation(settings, 'fr', 0.2)).toBe(false);
  expect(shouldOfferTranslation(settings, 'fr', 1)).toBe(true);
  expect(shouldOfferTranslation({ ...settings, mode: 'never' }, 'fr', 1)).toBe(false);
});
test('persists preferences via Outlook roaming settings and reports save failures', async () => {
  let stored: unknown;
  (globalThis as any).Office = { AsyncResultStatus: { Succeeded: 'succeeded' }, context: { roamingSettings: {
    get: () => stored, set: (_: string, value: unknown) => { stored = value; }, saveAsync: (cb: Function) => cb({ status: 'succeeded' })
  } } };
  await saveSettings({ mode: 'never', target: 'en', excluded: ['fr'] });
  expect(loadSettings()).toEqual({ mode: 'never', target: 'en', excluded: ['fr'] });
  (Office.context.roamingSettings as any).saveAsync = (cb: Function) => cb({ status: 'failed' });
  await expect(saveSettings({ mode: 'ask', target: 'en', excluded: [] })).rejects.toThrow('保存');
  expect(loadSettings().mode).toBe('never');
  delete (globalThis as any).Office;
});

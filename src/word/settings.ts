import { LANGUAGES } from '../shared/settings';

export interface WordSettings { target: string }
function normalizeSettings(input: unknown): WordSettings {
  const value = (input || {}) as Partial<WordSettings>;
  return { target: value.target && Object.hasOwnProperty.call(LANGUAGES, value.target) ? value.target : 'zh-Hans' };
}

const KEY = 'wordTranslation.preferences.v1';
// Document settings are supported by Word and persist with this document.
export function loadSettings(): WordSettings {
  return normalizeSettings(Office.context.document.settings.get(KEY));
}
export function saveSettings(value: WordSettings): Promise<void> {
  const storage = Office.context.document.settings;
  const previous = storage.get(KEY);
  storage.set(KEY, normalizeSettings(value));
  return new Promise((resolve, reject) => storage.saveAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
    else {
      if (previous === undefined || previous === null) storage.remove(KEY);
      else storage.set(KEY, previous);
      reject(new Error('保存设置失败，请重试。'));
    }
  }));
}

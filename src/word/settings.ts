import { normalizeSettings, Settings } from '../shared/settings';

const KEY = 'wordTranslation.preferences.v1';
// Document settings are supported by Word and persist with this document.
export function loadSettings(): Settings {
  return normalizeSettings(Office.context.document.settings.get(KEY));
}
export function saveSettings(value: Settings): Promise<void> {
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

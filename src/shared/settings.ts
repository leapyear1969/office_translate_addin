export const LANGUAGES: Record<string, string> = {
  'zh-Hans': '中文（简体）', 'zh-Hant': '中文（繁體）', en: '英语', ja: '日语', ko: '韩语',
  fr: '法语', de: '德语', es: '西班牙语', pt: '葡萄牙语', it: '意大利语', ru: '俄语',
  ar: '阿拉伯语', th: '泰语', vi: '越南语', id: '印度尼西亚语', ms: '马来语',
  nl: '荷兰语', pl: '波兰语', tr: '土耳其语', uk: '乌克兰语', hi: '印地语',
  sv: '瑞典语', da: '丹麦语', fi: '芬兰语', nb: '挪威语', cs: '捷克语', el: '希腊语',
  he: '希伯来语', hu: '匈牙利语', ro: '罗马尼亚语', bg: '保加利亚语',
};
export interface Settings { mode: 'always' | 'ask' | 'never'; target: string; excluded: string[] }
const KEY = 'mailTranslation.preferences.v1';
export function normalizeSettings(input: unknown): Settings {
  const value = (input || {}) as Partial<Settings>;
  return { mode: ['always', 'ask', 'never'].includes(value.mode || '') ? value.mode! : 'ask',
    target: value.target && Object.hasOwnProperty.call(LANGUAGES, value.target) ? value.target : 'zh-Hans',
    excluded: Array.isArray(value.excluded) ? [...new Set(value.excluded.filter(code => Object.hasOwnProperty.call(LANGUAGES, code)))] : [] };
}
export function loadSettings(): Settings { return normalizeSettings(Office.context.roamingSettings.get(KEY)); }
export function saveSettings(settings: Settings): Promise<void> {
  const previous = Office.context.roamingSettings.get(KEY);
  Office.context.roamingSettings.set(KEY, normalizeSettings(settings));
  return new Promise((resolve, reject) => Office.context.roamingSettings.saveAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
    else { Office.context.roamingSettings.set(KEY, previous); reject(new Error('保存设置失败，请重试。')); }
  }));
}
export function shouldOfferTranslation(settings: Settings, language: string, confidence: number): boolean {
  return settings.mode !== 'never' && confidence >= 0.7 && language !== settings.target && !settings.excluded.includes(language);
}

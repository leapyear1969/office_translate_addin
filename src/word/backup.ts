const KEY = 'wordTranslation.originalBody.v1';

export function originalBody(): string | undefined {
  const value = Office.context.document.settings.get(KEY);
  if (value == null) return undefined;
  if (typeof value !== 'string' || !value.trim()) throw new Error('原文备份无效，已取消操作。');
  return value;
}

export async function saveOriginalBody(ooxml: string): Promise<void> {
  if (originalBody() !== undefined) return;
  if (!ooxml.trim()) throw new Error('未能读取原文备份，已取消翻译。');
  const storage = Office.context.document.settings;
  try {
    storage.set(KEY, ooxml);
    await new Promise<void>((resolve, reject) => storage.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error('保存失败'));
    }));
  } catch {
    storage.remove(KEY);
    throw new Error('保存原文备份失败，未替换原文，请重试。');
  }
}

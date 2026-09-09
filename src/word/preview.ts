import { api, authenticate } from '../shared/api';
import { prepareOoxml } from './ooxml';

export class EmptySelectionError extends Error {
  constructor() {
    super('请先选中需要翻译的文字，或在原文框中输入内容。');
    this.name = 'EmptySelectionError';
  }
}

export interface PreviewRange {
  text: string;
  insert(text: string, shouldContinue: () => boolean): Promise<void>;
  release(): Promise<void>;
}

// Keep the original range even when focus or the document selection moves.
export async function capturePreview(scope: 'selection' | 'paragraph'): Promise<PreviewRange> {
  return Word.run(async context => {
    const selection = context.document.getSelection();
    const range = scope === 'paragraph' ? selection.paragraphs.getFirst().getRange() : selection;
    range.load('text,isEmpty');
    await context.sync();
    // A collapsed cursor has no OOXML to translate; some hosts throw when asked for it.
    if (scope === 'selection' && range.isEmpty) throw new EmptySelectionError();
    const original = range.text;
    let placeholderPlan: ReturnType<typeof prepareOoxml> | undefined;
    let text = original;
    if (!original.trim()) {
      const native = range.getOoxml();
      await context.sync();
      try { placeholderPlan = prepareOoxml(native.value); }
      catch { throw new Error(scope === 'selection' ? '请先选中需要翻译的文字，或选择可翻译的普通模板内容。' : '当前段落没有可安全翻译的文字。'); }
      text = placeholderPlan.paragraphs.map(html => new DOMParser().parseFromString(html, 'text/html').body.textContent || '').join('\n');
    }
    context.trackedObjects.add(range);
    await context.sync();
    let released = false;
    return {
      text,
      async insert(text, shouldContinue) {
        if (released) throw new Error('请重新选择翻译范围。');
        await Word.run(range, async ctx => {
          range.load('text');
          await ctx.sync();
          if (!shouldContinue()) throw new Error('已取消插入。');
          if (range.text !== original) throw new Error('原文范围已更改，请重新选择并翻译，避免覆盖您的编辑。');
          if (placeholderPlan) {
            const native = range.getOoxml();
            await ctx.sync();
            const current = prepareOoxml(native.value);
            if (current.sourceText !== placeholderPlan.sourceText || current.mapping !== placeholderPlan.mapping) {
              throw new Error('原文范围已更改，请重新选择并翻译，避免覆盖您的编辑。');
            }
            if (!shouldContinue()) throw new Error('已取消插入。');
          }
          range.insertText(text, Word.InsertLocation.replace);
          await ctx.sync();
        });
      },
      async release() {
        if (released) return;
        released = true;
        await Word.run(range, async ctx => { ctx.trackedObjects.remove(range); await ctx.sync(); });
      },
    };
  });
}

export async function translatePreview(text: string, target: string): Promise<string> {
  if (text.length > 20000) throw new Error('预览内容超过 20000 字符，请缩小范围。');
  const session = await authenticate();
  // Text is escaped before entering the existing HTML translation endpoint.
  const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const result = await api<{ html: string }>('/api/translate', session.token, { html: `<div>${escaped}</div>`, to: target });
  if (typeof result.html !== 'string' || result.html.length > 1000000) throw new Error('译文无效，请重试。');
  const parsed = new DOMParser().parseFromString(result.html, 'text/html');
  const translated = parsed.body.textContent || '';
  if (!translated.trim()) throw new Error('译文为空，请重试。');
  return translated;
}

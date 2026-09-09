import { api, authenticate } from '../shared/api';

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
    range.load('text');
    await context.sync();
    const original = range.text;
    if (!original.trim()) throw new Error(scope === 'selection' ? '请先选中需要翻译的文字。' : '当前段落没有可翻译的文字。');
    context.trackedObjects.add(range);
    await context.sync();
    let released = false;
    return {
      text: original,
      async insert(text, shouldContinue) {
        if (released) throw new Error('请重新选择翻译范围。');
        await Word.run(range, async ctx => {
          range.load('text');
          await ctx.sync();
          if (!shouldContinue()) throw new Error('已取消插入。');
          if (range.text !== original) throw new Error('原文范围已更改，请重新选择并翻译，避免覆盖您的编辑。');
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

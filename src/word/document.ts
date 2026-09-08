import { api, authenticate } from '../shared/api';

export type Scope = 'selection' | 'paragraph' | 'body';

function rangeFor(context: Word.RequestContext, scope: Scope): Word.Range {
  if (scope === 'body') return context.document.body.getRange();
  const selection = context.document.getSelection();
  return scope === 'paragraph' ? selection.paragraphs.getFirst().getRange() : selection;
}

export async function documentHtml(): Promise<string> {
  return Word.run(async context => {
    const html = context.document.body.getRange().getHtml();
    await context.sync();
    return html.value;
  });
}

export async function translateDocument(scope: Scope, target: string, expectedHtml?: string, shouldContinue = () => true): Promise<void> {
  await Word.run(async context => {
    // Capture and track the original range before authentication or network waits.
    const range = rangeFor(context, scope);
    context.trackedObjects.add(range);
    try {
      range.load('text');
      const html = range.getHtml();
      const original = range.getOoxml();
      await context.sync();
      if (!range.text.trim()) throw new Error(scope === 'selection' ? '请先选中需要翻译的文字。' : '当前范围没有可翻译的文字。');
      if (html.value.length > 1000000) throw new Error('文档内容过长，请选择较小范围分次翻译。');
      if (expectedHtml !== undefined && html.value !== expectedHtml) throw new Error('文档已更改，请重新点击翻译正文全文。');
      const session = await authenticate();
      const result = await api<{ html: string }>('/api/translate', session.token, { html: html.value, to: target });
      if (typeof result.html !== 'string' || !result.html.trim() || result.html.length > 1000000) throw new Error('译文无效或超过显示限制。');
      const current = range.getOoxml();
      await context.sync();
      if (!shouldContinue()) throw new Error('已取消自动翻译。');
      if (current.value !== original.value) throw new Error('翻译期间原文已更改，已取消替换，请重试。');
      range.insertHtml(result.html, Word.InsertLocation.replace);
      await context.sync();
    } finally {
      context.trackedObjects.remove(range);
      await context.sync();
    }
  });
}

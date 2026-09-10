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

async function syncPreview(context: Word.RequestContext, step: string): Promise<void> {
  try { await context.sync(); }
  catch (error) {
    const failure = error as { code?: string; message?: string; debugInfo?: { errorLocation?: string } };
    const location = failure.debugInfo?.errorLocation;
    throw new Error(`${step}失败：${failure.code || failure.message || 'Word 未知错误'}${location ? `（${location}）` : ''}。`);
  }
}

// Keep the original range even when focus or the document selection moves.
export async function capturePreview(scope: 'selection' | 'paragraph'): Promise<PreviewRange> {
  return Word.run(async context => {
    const selection = context.document.getSelection();
    // Keep the paragraph terminator (and its enclosing cell/control boundary).
    const range = scope === 'paragraph' ? selection.paragraphs.getFirst().getRange('Content') : selection;
    range.load('text,isEmpty');
    await context.sync();
    // A collapsed cursor has no OOXML to translate; some hosts throw when asked for it.
    if (scope === 'selection' && range.isEmpty) throw new EmptySelectionError();
    const original = range.text;
    let placeholderPlan: ReturnType<typeof prepareOoxml> | undefined;
    let placeholderControl: Word.ContentControl | undefined;
    let text = original;
    // Placeholder text is not consistently empty across Word hosts. A range
    // enclosing the control also has no parent, so inspect contained controls.
    const parent = range.parentContentControlOrNullObject;
    const contained = range.contentControls;
    parent.load('isNullObject,id');
    contained.load('items/id');
    await syncPreview(context, '读取模板内容控件');
    const candidates = [...(!parent.isNullObject ? [parent] : []), ...contained.items]
      .filter((control, index, all) => all.findIndex(other => other.id === control.id) === index);
    if (!original.trim() || candidates.length) {
      const native = range.getOoxml();
      const exports = candidates.map(control => ({ control, xml: control.getOoxml() }));
      await syncPreview(context, '读取模板占位内容');
      const hasPlaceholder = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml')
        .getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'showingPlcHdr').length > 0;
      const prompts = exports.filter(item => hasPlaceholder(item.xml.value));
      if (!original.trim() || prompts.length || hasPlaceholder(native.value)) {
        try { placeholderPlan = prepareOoxml(native.value); }
        catch { throw new Error(scope === 'selection' ? '请先选中需要翻译的文字，或选择可翻译的普通模板内容。' : '当前段落没有可安全翻译的文字。'); }
        if (!original.trim()) text = placeholderPlan.paragraphs.map(html => new DOMParser().parseFromString(html, 'text/html').body.textContent || '').join('\n');
        const matches = prompts.map(item => ({ ...item, plan: prepareOoxml(item.xml.value) }))
          .filter(item => item.plan.sourceText === placeholderPlan!.sourceText &&
            JSON.stringify(item.plan.paragraphs) === JSON.stringify(placeholderPlan!.paragraphs));
        if (prompts.length && matches.length !== 1) {
          throw new Error('当前范围仅包含模板占位内容的一部分，请选中完整占位内容后重新翻译。');
        }
        if (matches.length === 1) {
          placeholderControl = matches[0].control;
          placeholderPlan = matches[0].plan;
        } else if (hasPlaceholder(native.value)) {
          throw new Error('无法定位模板占位内容控件，请选中完整占位内容后重新翻译。');
        }
      }
    }
    context.trackedObjects.add(range);
    if (placeholderControl) context.trackedObjects.add(placeholderControl);
    await context.sync();
    let released = false;
    return {
      text,
      async insert(text, shouldContinue) {
        if (released) throw new Error('请重新选择翻译范围。');
        await Word.run(range, async ctx => {
          range.load('text');
          await syncPreview(ctx, '检查原文范围');
          if (!shouldContinue()) throw new Error('已取消插入。');
          if (range.text !== original) throw new Error('原文范围已更改，请重新选择并翻译，避免覆盖您的编辑。');
          if (placeholderPlan) {
            const native = placeholderControl ? placeholderControl.getOoxml() : range.getOoxml();
            await ctx.sync();
            const current = prepareOoxml(native.value);
            if (current.sourceText !== placeholderPlan.sourceText || current.mapping !== placeholderPlan.mapping) {
              throw new Error('原文范围已更改，请重新选择并翻译，避免覆盖您的编辑。');
            }
            if (!shouldContinue()) throw new Error('已取消插入。');
          }
          if (placeholderControl) placeholderControl.insertText(text, Word.InsertLocation.replace);
          else range.insertText(text, Word.InsertLocation.replace);
          await syncPreview(ctx, placeholderControl ? '替换模板控件原文' : '替换原文');
        });
      },
      async release() {
        if (released) return;
        released = true;
        await Word.run(range, async ctx => {
          ctx.trackedObjects.remove(range);
          if (placeholderControl) ctx.trackedObjects.remove(placeholderControl);
          await ctx.sync();
        });
      },
    };
  });
}

export async function translatePreview(text: string, target: string): Promise<string> {
  if (text.length > 20000) throw new Error('预览内容超过 20000 字符，请缩小范围。');
  const session = await authenticate(false);
  // Text is escaped before entering the existing HTML translation endpoint.
  const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const result = await api<{ html: string }>('/api/translate', session.token, { html: `<div>${escaped}</div>`, to: target });
  if (typeof result.html !== 'string' || result.html.length > 1000000) throw new Error('译文无效，请重试。');
  const parsed = new DOMParser().parseFromString(result.html, 'text/html');
  const translated = parsed.body.textContent || '';
  if (!translated.trim()) throw new Error('译文为空，请重试。');
  return translated;
}

import { stripNestedBackups } from './backup-ooxml';
import { isPictureDrawing } from './pictures';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const XML = 'http://www.w3.org/XML/1998/namespace';
const MAX_PACKAGE = 20000000;
const escapeHtml = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function parse(xml: string): XMLDocument {
  if (!xml || xml.length > MAX_PACKAGE || /<!DOCTYPE/i.test(xml)) throw new Error('文档结构为空、过大或不受支持，已取消全文翻译。');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('无法读取 Word 原生结构，已取消全文翻译。');
  return doc;
}

/** Keep the complete local package, including image binaries and relationships.
 * Only small, generated HTML paragraphs are sent to the translation service. */
export function prepareOoxml(xml: string) {
  const doc = parse(stripNestedBackups(xml));
  const main = Array.from(doc.getElementsByTagNameNS(PKG, 'part'))
    .find(part => part.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const body = main?.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('Word 未返回完整的正文结构，已取消全文翻译。');
  const segments: { nodes: Element[]; html: string; index: number }[] = [];
  const sourceParagraphs: string[] = [];
  const fieldParagraphs = new Set<Element>();
  let fieldDepth = 0;
  // Complex fields can span multiple paragraphs; their result paragraphs may
  // otherwise look like ordinary text. Keep the whole field result unchanged.
  function scanFields(element: Element, paragraph?: Element) {
    if (element.namespaceURI === W && element.localName === 'p') paragraph = element;
    if (fieldDepth && paragraph) fieldParagraphs.add(paragraph);
    if (element.namespaceURI === W && element.localName === 'fldChar') {
      const type = element.getAttributeNS(W, 'fldCharType');
      if (type === 'begin') fieldDepth++;
      if (paragraph) fieldParagraphs.add(paragraph);
      if (type === 'end') fieldDepth = Math.max(0, fieldDepth - 1);
    }
    for (const child of Array.from(element.children)) scanFields(child, paragraph);
  }
  scanFields(body);
  let skipped = 0;
  for (const paragraph of Array.from(body.getElementsByTagNameNS(W, 'p'))) {
    // Text-box paragraphs are considered separately from their anchor paragraph.
    const nodes = Array.from(paragraph.getElementsByTagNameNS(W, 't')).filter(node => {
      let parent = node.parentElement;
      while (parent && !(parent.namespaceURI === W && parent.localName === 'p')) parent = parent.parentElement;
      return parent === paragraph;
    });
    const index = sourceParagraphs.length;
    sourceParagraphs.push(nodes.map(node => node.textContent || '').join(''));
    if (!nodes.some(node => node.textContent?.trim())) continue;
    let unsafe = fieldParagraphs.has(paragraph);
    for (let parent = paragraph.parentElement; parent && parent !== body; parent = parent.parentElement) {
      if (parent.namespaceURI === W && ['ins', 'del', 'moveFrom', 'moveTo', 'sdt'].includes(parent.localName)) {
        // Ordinary template controls are safe, but bound or locked ones aren't.
        if (parent.localName !== 'sdt' || parent.getElementsByTagNameNS(W, 'dataBinding').length
          || parent.getElementsByTagNameNS(W, 'lock').length) unsafe = true;
      }
    }
    // Fields, revisions, hyperlinks and breaks can carry
    // semantics that cannot safely be mapped to reordered translated runs.
    const children = Array.from(paragraph.children);
    if (children.some(child => child.namespaceURI !== W || !['pPr', 'r'].includes(child.localName))) unsafe = true;
    const runs = children.filter(child => child.namespaceURI === W && child.localName === 'r');
    if (runs.some(run => Array.from(run.children).some(child => !isPictureDrawing(child)
      && (child.namespaceURI !== W || !['rPr', 't'].includes(child.localName))))) unsafe = true;
    if (unsafe) { skipped++; continue; }
    // Never move translated words across an image. Each side is an independent
    // segment, including pictures and text that share a single run.
    const groups: Element[][] = [[]];
    for (const run of runs) for (const child of Array.from(run.children)) {
      if (isPictureDrawing(child)) groups.push([]);
      else if (child.namespaceURI === W && child.localName === 't') groups[groups.length - 1].push(child);
    }
    const candidates = groups.filter(group => group.some(node => node.textContent?.trim())).map(group => ({
      nodes: group, index,
      html: '<p>' + group.map((node, i) => `<span id="r${i}">${escapeHtml(node.textContent || '')}</span>`).join('') + '</p>',
    }));
    if (candidates.some(segment => segment.html.length > 40000)) { skipped++; continue; }
    segments.push(...candidates);
  }
  if (!segments.length) throw new Error(`没有可安全翻译的正文段落${skipped ? `，${skipped} 个复杂段落已保留原文` : ''}。`);
  if (segments.reduce((total, segment) => total + segment.html.length, 0) > 1000000) throw new Error('文档文字过长，请分次翻译。');
  return {
    paragraphs: segments.map(segment => segment.html),
    sourceText: JSON.stringify(sourceParagraphs),
    mapping: JSON.stringify(segments.map(segment => [segment.index, segment.nodes.map(node => node.textContent || '')])),
    apply(translations: unknown) {
      if (!Array.isArray(translations) || translations.length !== segments.length
        || translations.some(value => typeof value !== 'string')
        || translations.join('').length > 1000000) throw new Error('译文无效或不完整，已取消全文翻译。');
      let translated = 0;
      const skippedSegments = new Set<number>();
      segments.forEach((segment, index) => {
        const result = new DOMParser().parseFromString(translations[index], 'text/html');
        const paragraph = result.body.children[0];
        const spans = paragraph ? Array.from(paragraph.children) : [];
        // Fail closed for a paragraph if Translator drops/reorders markers or
        // inserts unexpected markup. Never guess where formatted text belongs.
        if (result.body.children.length !== 1 || paragraph?.tagName !== 'P'
          || spans.length !== segment.nodes.length
          || Array.from(result.body.childNodes).some(node => node !== paragraph && node.textContent?.trim())
          || Array.from(paragraph.childNodes).some(node => node.nodeType !== 1 && node.textContent?.trim())
          || spans.some((span, i) => span.tagName !== 'SPAN' || span.id !== `r${i}` || span.children.length
            || ((segment.nodes[i].textContent || '').trim() && !span.textContent?.trim()))) {
          skippedSegments.add(segment.index); return;
        }
        spans.forEach((span, i) => {
          segment.nodes[i].textContent = span.textContent;
          segment.nodes[i].setAttributeNS(XML, 'xml:space', 'preserve');
        });
        translated++;
      });
      if (!translated) throw new Error('译文格式标记无法安全对应原文，未修改正文。');
      const ooxml = new XMLSerializer().serializeToString(doc);
      if (ooxml.length > MAX_PACKAGE) throw new Error('译文文档结构过大，已取消全文翻译。');
      return { ooxml, skippedParagraphs: skipped + skippedSegments.size };
    },
  };
}

/** Exports aren't revision tokens. Reapply validated translations to the latest
 * package instead of comparing volatile IDs or inserting an older layout. */
export function rebaseOoxml(source: ReturnType<typeof prepareOoxml>, currentXml: string, translations: unknown) {
  const current = prepareOoxml(currentXml);
  if (current.sourceText !== source.sourceText) {
    throw new Error('翻译期间正文文字已更改，已取消替换，请重试。');
  }
  if (current.mapping !== source.mapping) {
    throw new Error('翻译期间段落的文字分段或可翻译范围发生变化，无法安全对应译文，已取消替换，请重试。');
  }
  return current.apply(translations);
}

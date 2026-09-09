import { isPictureDrawing } from './pictures';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
export const WEB_LIMITS = { characters: 20000, paragraphs: 500, packageBytes: 5000000 };
const guidance = '请先备份文档，再使用桌面 Word 客户端打开并翻译。';
function blocked(reason: string): never {
  throw new Error(`网页版已停止翻译：${reason}。${guidance}`);
}
export function isWordOnline(): boolean {
  return Office.context.platform !== undefined && Office.context.platform === Office.PlatformType.OfficeOnline;
}

// Only explicitly supported content and formatting are admitted. Package metadata
// (styles, relationships, etc.) is not itself document content.
const children: Record<string, string[]> = {
  body: ['p', 'tbl', 'sectPr'], p: ['pPr', 'r'], r: ['rPr', 't', 'tab', 'br', 'cr'], t: [], tab: [], br: [], cr: [],
  pPr: ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'spacing', 'ind', 'jc', 'outlineLvl', 'rPr', 'tabs', 'contextualSpacing', 'bidi', 'cnfStyle'],
  rPr: ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'vertAlign', 'rtl', 'cs', 'lang', 'noProof', 'shd'],
  tabs: ['tab'], sectPr: ['pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'cols', 'docGrid'],
  pgBorders: ['top', 'left', 'bottom', 'right'],
  tbl: ['tblPr', 'tblGrid', 'tr'], tblGrid: ['gridCol'],
  tblPr: ['tblStyle', 'tblW', 'jc', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblCellSpacing', 'tblLook', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblCaption', 'tblDescription'],
  tr: ['tblPrEx', 'trPr', 'tc'],
  tblPrEx: ['tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook'],
  trPr: ['cnfStyle', 'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'gridBefore', 'gridAfter'],
  tc: ['tcPr', 'p'], tcPr: ['cnfStyle', 'tcW', 'gridSpan', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'],
  tblBorders: ['top', 'left', 'bottom', 'right', 'start', 'end', 'insideH', 'insideV'],
  tcBorders: ['top', 'left', 'bottom', 'right', 'start', 'end', 'insideH', 'insideV', 'tl2br', 'tr2bl'],
  tblCellMar: ['top', 'left', 'bottom', 'right', 'start', 'end'], tcMar: ['top', 'left', 'bottom', 'right', 'start', 'end'],
};
// Check header/footer structure even when text is empty. Their ordinary tables
// are preserved, while the existing translator only translates body content.
function inspectHeaderFooter(root: Element, location: string, separator = false): void {
  function visit(node: Element) {
    if (node.namespaceURI !== W) blocked(`${location}包含尚未支持的结构`);
    if (separator && node.localName === 't' && node.textContent?.trim()) blocked(`${location}包含文字内容`);
    const allowed = ['hdr', 'ftr', 'body'].includes(node.localName) ? ['p', 'tbl']
      : ['footnote', 'endnote'].includes(node.localName) ? ['p']
      : separator && node.localName === 'r' ? [...children.r, 'separator', 'continuationSeparator'] : (children[node.localName] || []);
    for (const child of Array.from(node.children)) {
      if (!separator && node.localName === 'r' && isPictureDrawing(child)) continue;
      if (!allowed.includes(child.localName)) blocked(`${location}包含${labels[child.localName] || '尚未支持的结构或格式'}`);
      visit(child);
    }
  }
  visit(root);
}
const labels: Record<string, string> = {
  tbl: '表格', drawing: '图片或绘图', pict: '图片或文本框', object: '嵌入对象',
  hyperlink: '超链接', fldSimple: '域或目录', fldChar: '域或目录', instrText: '域或目录',
  ins: '修订', del: '修订', sdt: '内容控件', numPr: '自动编号或列表',
  commentRangeStart: '批注', footnoteReference: '脚注', endnoteReference: '尾注',
  headerReference: '页眉', footerReference: '页脚',
};

function inspectTables(doc: XMLDocument): void {
  const direct = (node: Element, name: string) => Array.from(node.children)
    .filter(child => child.namespaceURI === W && child.localName === name);
  for (const table of Array.from(doc.getElementsByTagNameNS(W, 'tbl'))) {
    if (table.getElementsByTagNameNS(W, 'tbl').length) blocked('文档包含嵌套表格');
    if (table.getElementsByTagNameNS(W, 'vMerge').length || table.getElementsByTagNameNS(W, 'hMerge').length) blocked('表格包含合并单元格');
    for (const span of Array.from(table.getElementsByTagNameNS(W, 'gridSpan'))) {
      if (span.getAttributeNS(W, 'val') !== '1') blocked('表格包含合并单元格或无法确认的跨列结构');
    }
    for (const name of ['gridBefore', 'gridAfter']) {
      if (Array.from(table.getElementsByTagNameNS(W, name)).some(node => node.getAttributeNS(W, 'val') !== '0')) {
        blocked('表格包含不完整的行或缺失单元格');
      }
    }
    const rows = direct(table, 'tr');
    const counts = rows.map(row => direct(row, 'tc').length);
    const grids = direct(table, 'tblGrid');
    if (!rows.length || !counts[0] || counts.some(count => count !== counts[0])
      || grids.length > 1 || (grids.length === 1 && direct(grids[0], 'gridCol').length !== counts[0])) {
      blocked('表格行列结构不规则或无法确认');
    }
  }
}

export function inspectWebOoxml(xml: string, options: { headerFooter?: boolean; headerFooterChecked?: boolean } = {}): void {
  let bytes = 0;
  for (const char of xml) {
    const code = char.codePointAt(0)!;
    bytes += code <= 0x7f ? 1 : code <= 0x7ff ? 2 : code <= 0xffff ? 3 : 4;
    if (bytes > WEB_LIMITS.packageBytes) blocked('文档结构数据超过 5 MB');
  }
  if (!xml.trim() || /<!DOCTYPE/i.test(xml)) blocked('无法确认文档结构');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) blocked('无法读取文档结构');
  inspectTables(doc);
  const parts = Array.from(doc.getElementsByTagNameNS(PKG, 'part'));
  const mains = parts.filter(part => part.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const bodies = mains[0]?.getElementsByTagNameNS(W, 'body');
  if (mains.length !== 1 || bodies?.length !== 1) blocked('无法确认完整正文结构');
  const mainDocument = bodies![0].parentElement;
  if (mainDocument?.namespaceURI !== W || mainDocument.localName !== 'document'
    || Array.from(mainDocument.children).some(child => child !== bodies![0])) blocked('文档包含尚未支持的文档级结构');
  for (const part of parts) {
    const name = part.getAttributeNS(PKG, 'name') || '';
    const binaryCount = part.getElementsByTagNameNS(PKG, 'binaryData').length;
    if (/\/media\//i.test(name) && /^image\//i.test(part.getAttributeNS(PKG, 'contentType') || '')
      && binaryCount === 1 && !part.getElementsByTagNameNS(PKG, 'xmlData').length) continue;
    if (/\/(media|embeddings)\//i.test(name) || binaryCount) blocked('文档包含尚未支持的图片或嵌入资源');
  }
  for (const part of parts) {
    const name = part.getAttributeNS(PKG, 'name') || '';
    if (/^\/word\/(header|footer)[^/]*\.xml$/i.test(name)) {
      const header = /^\/word\/header/i.test(name);
      const roots = part.getElementsByTagNameNS(W, header ? 'hdr' : 'ftr');
      if (roots.length !== 1) blocked('无法确认页眉或页脚结构');
      inspectHeaderFooter(roots[0], header ? '页眉' : '页脚');
    } else if (/^\/word\/(footnotes|endnotes|comments)[^/]*\.xml$/i.test(name)) {
      const kind = /^\/word\/footnotes\.xml$/i.test(name) ? 'footnotes'
        : /^\/word\/endnotes\.xml$/i.test(name) ? 'endnotes'
        : /^\/word\/comments\.xml$/i.test(name) ? 'comments' : '';
      const roots = kind ? part.getElementsByTagNameNS(W, kind) : [];
      if (roots.length !== 1) blocked('无法确认脚注、尾注或批注结构');
      for (const note of Array.from(roots[0].children)) {
        if (kind === 'comments' || note.namespaceURI !== W || note.localName !== kind.slice(0, -1)
          || !['separator', 'continuationSeparator'].includes(note.getAttributeNS(W, 'type') || '')) {
          blocked('文档包含脚注、尾注或批注内容');
        }
        inspectHeaderFooter(note, '脚注或尾注分隔符', true);
      }
    }
  }
  const body = bodies![0];
  if (options.headerFooter) {
    // A header/footer range export is represented as a document body by Word.
    // Permit its export-only final section properties and ordinary content.
    for (const child of Array.from(body.children)) {
      if (child.namespaceURI === W && child.localName === 'sectPr') continue;
      if (child.namespaceURI !== W || !['p', 'tbl'].includes(child.localName)) blocked(`页眉或页脚包含${labels[child.localName] || '尚未支持的结构'}`);
      inspectHeaderFooter(child, '页眉或页脚');
    }
    return;
  }
  if (body.getElementsByTagNameNS(W, 'p').length > WEB_LIMITS.paragraphs) blocked('文档超过 500 个段落');
  const text = Array.from(body.getElementsByTagNameNS(W, 't')).map(node => node.textContent || '').join('');
  if (text.length > WEB_LIMITS.characters) blocked('文档超过 20000 个文字字符');
  function visit(node: Element) {
    if (node.localName === 'cols' && Number(node.getAttributeNS(W, 'num') || '1') !== 1) blocked('文档包含分栏');
    if (node.localName === 'sectPr' && node.parentElement !== body) blocked('文档包含分节');
    for (const child of Array.from(node.children)) {
      if (node.localName === 'r' && isPictureDrawing(child)) continue;
      if (options.headerFooterChecked && node.localName === 'sectPr' && child.namespaceURI === W
        && ['headerReference', 'footerReference'].includes(child.localName) && !child.children.length) continue;
      if (child.namespaceURI !== W || !(children[node.localName] || []).includes(child.localName)) {
        blocked(`文档包含${labels[child.localName] || '尚未支持的结构或格式'}`);
      }
      visit(child);
    }
  }
  visit(body);
}

export async function checkWebDocument(context: Word.RequestContext): Promise<void> {
  if (!isWordOnline()) return;
  try {
    const body = context.document.body.getRange();
    body.load('text');
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();
    if (body.text.length > WEB_LIMITS.characters) blocked('文档超过 20000 个文字字符');
    if (sections.items.length !== 1) blocked('文档包含分节或无法确认章节结构');
    // Body exports need not include every header/footer. Inspect these separately.
    const extras = (['Primary', 'FirstPage', 'EvenPages'] as const).flatMap(type =>
      [sections.items[0].getHeader(type), sections.items[0].getFooter(type)]);
    // Office.js is a fast rejection pass, never proof that no merges exist.
    // Some merges leave all rows uniform; OOXML must always be checked next.
    const tables = [context.document.body, ...extras].map(part => part.tables);
    tables.forEach(collection => collection.load('items/isUniform,items/nestingLevel'));
    await context.sync();
    for (const collection of tables) for (const table of collection.items) {
      if (table.nestingLevel > 1) blocked('文档包含嵌套表格');
      if (table.isUniform === false) blocked('表格行列不一致，可能包含合并单元格');
      if (table.isUniform !== true || table.nestingLevel !== 1) blocked('无法确认表格结构');
    }
    const xml = body.getOoxml();
    const extraXml = extras.map(extra => extra.getOoxml());
    await context.sync();
    extraXml.forEach(item => inspectWebOoxml(item.value, { headerFooter: true }));
    inspectWebOoxml(xml.value, { headerFooterChecked: true });
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('网页版已停止翻译：')) throw error;
    blocked('无法完成文档兼容性检查');
  }
}

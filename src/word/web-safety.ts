const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
export const WEB_LIMITS = { characters: 20000, paragraphs: 500, packageBytes: 2000000 };
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
  body: ['p', 'sectPr'], p: ['pPr', 'r'], r: ['rPr', 't', 'tab', 'br', 'cr'], t: [], tab: [], br: [], cr: [],
  pPr: ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'spacing', 'ind', 'jc', 'outlineLvl', 'rPr', 'tabs', 'contextualSpacing'],
  rPr: ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'vertAlign', 'rtl', 'cs', 'lang', 'noProof', 'shd'],
  tabs: ['tab'], sectPr: ['pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'cols', 'docGrid'],
  pgBorders: ['top', 'left', 'bottom', 'right'],
};
const labels: Record<string, string> = {
  tbl: '表格', drawing: '图片或绘图', pict: '图片或文本框', object: '嵌入对象',
  hyperlink: '超链接', fldSimple: '域或目录', fldChar: '域或目录', instrText: '域或目录',
  ins: '修订', del: '修订', sdt: '内容控件', numPr: '自动编号或列表',
  commentRangeStart: '批注', footnoteReference: '脚注', endnoteReference: '尾注',
  headerReference: '页眉', footerReference: '页脚',
};

export function inspectWebOoxml(xml: string): void {
  let bytes = 0;
  for (const char of xml) {
    const code = char.codePointAt(0)!;
    bytes += code <= 0x7f ? 1 : code <= 0x7ff ? 2 : code <= 0xffff ? 3 : 4;
    if (bytes > WEB_LIMITS.packageBytes) blocked('文档结构数据超过 2 MB');
  }
  if (!xml.trim() || /<!DOCTYPE/i.test(xml)) blocked('无法确认文档结构');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) blocked('无法读取文档结构');
  const parts = Array.from(doc.getElementsByTagNameNS(PKG, 'part'));
  const mains = parts.filter(part => part.getAttributeNS(PKG, 'name') === '/word/document.xml');
  const bodies = mains[0]?.getElementsByTagNameNS(W, 'body');
  if (mains.length !== 1 || bodies?.length !== 1) blocked('无法确认完整正文结构');
  const mainDocument = bodies![0].parentElement;
  if (mainDocument?.namespaceURI !== W || mainDocument.localName !== 'document'
    || Array.from(mainDocument.children).some(child => child !== bodies![0])) blocked('文档包含尚未支持的文档级结构');
  if (parts.some(part => /\/(media|embeddings)\//i.test(part.getAttributeNS(PKG, 'name') || '')
    || part.getElementsByTagNameNS(PKG, 'binaryData').length)) blocked('文档包含图片或嵌入资源');
  if (parts.some(part => /\/word\/(header|footer|footnotes|endnotes|comments)/i.test(part.getAttributeNS(PKG, 'name') || ''))) blocked('文档包含页眉、页脚、注释等附加结构');
  const body = bodies![0];
  if (body.getElementsByTagNameNS(W, 'p').length > WEB_LIMITS.paragraphs) blocked('文档超过 500 个段落');
  const text = Array.from(body.getElementsByTagNameNS(W, 't')).map(node => node.textContent || '').join('');
  if (text.length > WEB_LIMITS.characters) blocked('文档超过 20000 个文字字符');
  function visit(node: Element) {
    if (node.localName === 'cols' && Number(node.getAttributeNS(W, 'num') || '1') !== 1) blocked('文档包含分栏');
    if (node.localName === 'sectPr' && node.parentElement !== body) blocked('文档包含分节');
    for (const child of Array.from(node.children)) {
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
    const xml = body.getOoxml();
    // Body exports need not include every header/footer. Inspect these separately.
    const extras = (['Primary', 'FirstPage', 'EvenPages'] as const).flatMap(type =>
      [sections.items[0].getHeader(type), sections.items[0].getFooter(type)]);
    extras.forEach(extra => extra.load('text'));
    const extraXml = extras.map(extra => extra.getOoxml());
    await context.sync();
    inspectWebOoxml(xml.value);
    if (extras.some(extra => extra.text.trim())) blocked('文档包含页眉或页脚内容');
    extraXml.forEach(item => inspectWebOoxml(item.value));
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('网页版已停止翻译：')) throw error;
    blocked('无法完成文档兼容性检查');
  }
}

/** @jest-environment jsdom */
import { inspectWebOoxml, WEB_LIMITS } from './web-safety';

export const simplePackage = (content = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>') =>
  `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${content}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

test('allows plain text, heading styles and basic formatting', () => {
  expect(() => inspectWebOoxml(simplePackage('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:after="100"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>标题</w:t></w:r></w:p>'))).not.toThrow();
});
test.each(['tbl', 'drawing', 'pict', 'object', 'hyperlink', 'fldChar', 'ins', 'del', 'sdt', 'numPr', 'commentRangeStart', 'footnoteReference', 'unknown'])('blocks %s and directs to desktop', element => {
  expect(() => inspectWebOoxml(simplePackage(`<w:p><w:${element}/></w:p>`))).toThrow('桌面 Word');
});
test.each(['<bad/>', '', '<!DOCTYPE x><x/>', '<broken'])('fails closed for invalid exports', xml => {
  expect(() => inspectWebOoxml(xml)).toThrow('网页版已停止');
});
test('enforces character, paragraph and byte limits', () => {
  expect(() => inspectWebOoxml(simplePackage(`<w:p><w:r><w:t>${'中'.repeat(WEB_LIMITS.characters + 1)}</w:t></w:r></w:p>`))).toThrow('20000');
  expect(() => inspectWebOoxml(simplePackage('<w:p/>'.repeat(501)))).toThrow('500');
  expect(() => inspectWebOoxml('中'.repeat(666667))).toThrow('2 MB');
});
test('rejects columns, section breaks and foreign namespace content', () => {
  for (const content of ['<w:sectPr><w:cols w:num="2"/></w:sectPr>', '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>', '<w:p><math xmlns="urn:math"/></w:p>']) {
    expect(() => inspectWebOoxml(simplePackage(content))).toThrow('桌面 Word');
  }
});
test.each(['/word/media/image.png', '/word/comments.xml', '/word/header1.xml'])('rejects additional package content %s', name => {
  const xml = simplePackage().replace('</pkg:package>', `<pkg:part pkg:name="${name}"><pkg:xmlData/></pkg:part></pkg:package>`);
  expect(() => inspectWebOoxml(xml)).toThrow('桌面 Word');
});

const withPart = (name: string, content: string, xml = simplePackage()) => xml.replace('</pkg:package>',
  `<pkg:part pkg:name="/word/${name}.xml"><pkg:xmlData>${content}</pkg:xmlData></pkg:part></pkg:package>`);

test.each(['header', 'footer'])('allows a genuinely empty %s part', name => {
  const tag = name === 'header' ? 'hdr' : 'ftr';
  expect(() => inspectWebOoxml(withPart(name, `<w:${tag}><w:p><w:pPr><w:pStyle w:val="Header"/><w:bidi w:val="0"/></w:pPr></w:p></w:${tag}>`))).not.toThrow();
});
test.each(['<w:p><w:r><w:drawing/></w:r></w:p>', '<w:p><w:r><w:fldChar/></w:r></w:p>'])('rejects complex header contents even with no visible text: %s', contents => {
  expect(() => inspectWebOoxml(withPart('header', `<w:hdr>${contents}</w:hdr>`))).toThrow('页眉');
  expect(() => inspectWebOoxml(simplePackage(contents), { headerFooter: true })).toThrow('页眉');
});
test('Hello Jason with six empty header/footer tables is allowed', () => {
  let xml = simplePackage('<w:p><w:r><w:rPr/><w:t>Hello Jason</w:t></w:r></w:p>');
  const table = '<w:tbl><w:tr>' + '<w:tc><w:p/></w:tc>'.repeat(3) + '</w:tr></w:tbl><w:p/>';
  for (const name of ['header', 'footer']) for (const suffix of ['', '2', '3']) {
    const tag = name === 'header' ? 'hdr' : 'ftr';
    xml = withPart(name + suffix, `<w:${tag}>${table}</w:${tag}>`, xml);
  }
  expect(() => inspectWebOoxml(xml)).not.toThrow();
});

const cell = (content = '<w:p/>', properties = '') => `<w:tc><w:tcPr>${properties}</w:tcPr>${content}</w:tc>`;
const table = (rows: string[], grid = '') => `<w:tbl>${grid}${rows.map(row => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
test.each(['', '<w:p><w:r><w:t>Cell text</w:t></w:r></w:p>'])('allows rectangular tables, including empty cells: %s', content => {
  const xml = table([cell(content || '<w:p/>') + cell(), cell() + cell()]);
  expect(() => inspectWebOoxml(simplePackage(xml))).not.toThrow();
  expect(() => inspectWebOoxml(withPart('header', `<w:hdr>${xml}</w:hdr>`))).not.toThrow();
});
test.each(['<w:gridSpan w:val="2"/>', '<w:vMerge w:val="restart"/>', '<w:vMerge/>', '<w:hMerge w:val="restart"/>', '<w:hMerge/>'])('rejects all merge encodings %s even when every row has the same cell count', merge => {
  const xml = table([cell('<w:p/>', merge), cell('<w:p/>', merge)]);
  expect(() => inspectWebOoxml(simplePackage(xml))).toThrow('合并');
  expect(() => inspectWebOoxml(withPart('footer', `<w:ftr>${xml}</w:ftr>`))).toThrow('合并');
});
test('rejects nested tables even with regular rows', () => {
  const xml = table([cell(table([cell()]))]);
  expect(() => inspectWebOoxml(simplePackage(xml))).toThrow('嵌套');
  expect(() => inspectWebOoxml(simplePackage(xml), { headerFooter: true })).toThrow('嵌套');
});
test('rejects ragged rows, grid mismatches and omitted cells', () => {
  for (const xml of [table([cell(), cell() + cell()]), table([cell()], '<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>'),
    table(['<w:trPr><w:gridBefore w:val="1"/></w:trPr>' + cell()])]) {
    expect(() => inspectWebOoxml(simplePackage(xml))).toThrow('表格');
  }
});
test('ordinary tables do not bypass other content restrictions', () => {
  expect(() => inspectWebOoxml(simplePackage(table([cell('<w:p><w:r><w:drawing/></w:r></w:p>')])))).toThrow('图片');
  expect(() => inspectWebOoxml(simplePackage(table([cell('<w:p><w:r><w:t>' + 'a'.repeat(20001) + '</w:t></w:r></w:p>')])))).toThrow('20000');
});
test('header/footer references require independently checked contents', () => {
  const xml = simplePackage('<w:p/><w:sectPr><w:headerReference/></w:sectPr>');
  expect(() => inspectWebOoxml(xml)).toThrow('页眉');
  expect(() => inspectWebOoxml(xml, { headerFooterChecked: true })).not.toThrow();
});
test('allows empty comments and default note separators but blocks actual notes', () => {
  expect(() => inspectWebOoxml(withPart('comments', '<w:comments/>'))).not.toThrow();
  for (const kind of ['footnote', 'endnote']) {
    const xml = withPart(kind + 's', `<w:${kind}s><w:${kind} w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}></w:${kind}s>`);
    expect(() => inspectWebOoxml(xml)).not.toThrow();
    expect(() => inspectWebOoxml(xml.replace('w:type="separator"', 'w:type="normal"'))).toThrow('内容');
    expect(() => inspectWebOoxml(xml.replace('<w:separator/>', '<w:drawing/>'))).toThrow('图片');
  }
});

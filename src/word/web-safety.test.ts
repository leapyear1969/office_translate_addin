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

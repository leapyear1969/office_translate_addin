/** @jest-environment jsdom */
import { prepareOoxml, ooxmlStructure } from './ooxml';
import { wordPackage } from './test-fixtures/package';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');
const serialize = (node: Node) => new XMLSerializer().serializeToString(node);
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

test('changes only text nodes while preserving newsletter layout, mixed fonts, images and package parts', () => {
  const original = wordPackage(`<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:shd w:fill="E6EBEF"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Title"/><w:spacing w:after="240"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="48"/></w:rPr><w:t>Modern </w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="0078D4"/></w:rPr><w:t>living</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:drawing><anchor xmlns="urn:drawing" position="123" image="rId1"/></w:drawing></w:r></w:p>`);
  const plan = prepareOoxml(original);
  expect(plan.paragraphs).toEqual(['<p><span id="r0">Modern </span><span id="r1">living</span></p>']);
  expect(plan.paragraphs.join('')).not.toMatch(/binaryData|aW1hZ2U|tblW|drawing/);
  const result = plan.apply(['<p><span id="r0">现代</span><span id="r1">生活 &amp; 家居</span></p>']);
  expect(result.skippedParagraphs).toBe(0);
  const before = parse(original), after = parse(result.ooxml);
  expect(Array.from(after.getElementsByTagNameNS(W, 't')).map(node => node.textContent)).toEqual(['现代', '生活 & 家居']);
  // Erase only text and its whitespace-preservation attribute, then compare ALL
  // remaining structure, including table geometry, fonts and image references.
  for (const doc of [before, after]) for (const node of Array.from(doc.getElementsByTagNameNS(W, 't'))) {
    node.textContent = ''; node.removeAttributeNS('http://www.w3.org/XML/1998/namespace', 'space');
  }
  expect(serialize(after)).toBe(serialize(before));
});

test('translates text-box paragraphs without touching their drawing anchors', () => {
  const plan = prepareOoxml(wordPackage(`<w:p><w:r><w:pict><shape xmlns="urn:vml" style="position:absolute;width:100pt"><w:txbxContent>${paragraph('Box')}</w:txbxContent></shape></w:pict></w:r></w:p>`));
  expect(plan.paragraphs).toEqual(['<p><span id="r0">Box</span></p>']);
  expect(plan.apply(['<p><span id="r0">文本框</span></p>']).ooxml).toContain('position:absolute;width:100pt');
});

test('keeps complex fields, revisions and bound controls untouched and reports skipped paragraphs', () => {
  const fields = '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>Date</w:t></w:r></w:fldSimple></w:p>';
  const revision = '<w:p><w:ins><w:r><w:t>Edited</w:t></w:r></w:ins></w:p>';
  const bound = `<w:sdt><w:sdtPr><w:dataBinding w:xpath="/value"/></w:sdtPr><w:sdtContent>${paragraph('Bound')}</w:sdtContent></w:sdt>`;
  const plan = prepareOoxml(wordPackage(fields + revision + bound + paragraph('Safe')));
  expect(plan.paragraphs).toEqual(['<p><span id="r0">Safe</span></p>']);
  const result = plan.apply(['<p><span id="r0">安全</span></p>']);
  expect(result.skippedParagraphs).toBe(3);
  expect(result.ooxml).toContain('>Date<');
  expect(result.ooxml).toContain('>Edited<');
  expect(result.ooxml).toContain('>Bound<');
});

test.each(['<p>Lost markers</p>', '<p><span id="r1">wrong</span></p>', '<p><span id="r0"></span></p>',
  '<p><span id="r0"><b>extra markup</b></span></p>', '<p><span id="r0">ok</span>unmapped</p>'])('skips unsafe translation markup: %s', invalid => {
  const plan = prepareOoxml(wordPackage(paragraph('First') + paragraph('Second')));
  const result = plan.apply([invalid, '<p><span id="r0">第二</span></p>']);
  expect(result.skippedParagraphs).toBe(1);
  expect(result.ooxml).toContain('>First<');
  expect(result.ooxml).toContain('>第二<');
});

test.each([null, [], [42], ['<p>bad markers</p>']])('rejects invalid/incomplete results before any write: %j', value => {
  expect(() => prepareOoxml(wordPackage()).apply(value)).toThrow();
});

test.each(['', '<broken>', '<html/>', '<!DOCTYPE x><x/>'])('rejects unsupported or malformed OOXML: %s', xml => {
  expect(() => prepareOoxml(xml)).toThrow();
});

test('skips excessively long paragraphs without splitting their formatting markers', () => {
  const plan = prepareOoxml(wordPackage(paragraph('x'.repeat(40000)) + paragraph('Safe')));
  expect(plan.paragraphs).toHaveLength(1);
  expect(plan.apply(['<p><span id="r0">安全</span></p>']).skippedParagraphs).toBe(1);
});

test('does not translate field results spanning multiple paragraphs', () => {
  const plan = prepareOoxml(wordPackage('<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC</w:instrText></w:r></w:p>'
    + paragraph('Field result') + '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + paragraph('Normal')));
  expect(plan.paragraphs).toEqual(['<p><span id="r0">Normal</span></p>']);
  expect(plan.apply(['<p><span id="r0">正常</span></p>']).ooxml).toContain('>Field result<');
});

test('structure comparison ignores export whitespace and settings but detects layout and image edits', () => {
  const original = wordPackage();
  const metadata = original.replace('<w:p>', '<w:p w:rsidR="123">').replace(/></g, '>\n<')
    .replace('</pkg:package>', '<pkg:part pkg:name="/word/settings.xml"><pkg:xmlData><settings/></pkg:xmlData></pkg:part></pkg:package>');
  expect(ooxmlStructure(metadata)).toBe(ooxmlStructure(original));
  expect(ooxmlStructure(original.replace('11906', '15000'))).not.toBe(ooxmlStructure(original));
  expect(ooxmlStructure(original.replace('aW1hZ2U=', 'bmV3'))).not.toBe(ooxmlStructure(original));
});

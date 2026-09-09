export const listParagraph = (numId = 1, level = 0, text = 'Original') =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

export const withNumbering = (xml: string) => xml.replace('</pkg:package>',
  `<pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData><w:numbering>
    <w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl>
    </w:abstractNum>
    <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
    <w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
    <w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num>
  </w:numbering></pkg:xmlData></pkg:part></pkg:package>`);

export const picture = (layout: 'inline' | 'anchor' = 'inline') => `<w:drawing
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <wp:${layout}><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>
 <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
 <pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>
 <pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
 <pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
 </a:graphicData></a:graphic></wp:${layout}></w:drawing>`;

export const withPictureParts = (xml: string) => xml.replace('</pkg:package>', `
 <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png"><pkg:binaryData>aW1hZ2U=</pkg:binaryData></pkg:part>
 <pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData>
 <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
 </Relationships></pkg:xmlData></pkg:part></pkg:package>`);

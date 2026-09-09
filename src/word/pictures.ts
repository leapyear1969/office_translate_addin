const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

/** Ordinary DrawingML pictures only; charts, SmartArt and text boxes are not pictures. */
export function isPictureDrawing(node: Element): boolean {
  if (node.namespaceURI !== W || node.localName !== 'drawing' || node.children.length !== 1) return false;
  const frame = node.children[0];
  if (frame.namespaceURI !== WP || !['inline', 'anchor'].includes(frame.localName)) return false;
  const data = node.getElementsByTagNameNS(A, 'graphicData');
  return data.length === 1 && data[0].getAttribute('uri') === PIC
    && data[0].children.length === 1 && data[0].children[0].namespaceURI === PIC
    && data[0].children[0].localName === 'pic'
    && node.getElementsByTagNameNS(A, 'blip').length === 1
    && !node.getElementsByTagNameNS(W, '*').length
    && !node.getElementsByTagNameNS(A, 't').length;
}

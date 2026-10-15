// Construcción de paquetes OOXML (Flat OPC) para insertOoxml. Helpers puros.

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:w10="urn:schemas-microsoft-com:office:word"',
].join(" ");

/**
 * Acepta un paquete completo (<pkg:package>), un <w:document> o solo el contenido del cuerpo
 * (<w:p>, <w:tbl>, <w:sdt>…) y devuelve un paquete listo para insertOoxml.
 */
export function wrapOoxml(xml: string): string {
  const source = xml.trim().replace(/^<\?xml[^>]*\?>\s*/, "");
  if (source.includes("<pkg:package")) return source;
  let document: string;
  if (source.includes("<w:document")) {
    document = source;
  } else {
    const body = source.startsWith("<w:body") ? source : `<w:body>${source}</w:body>`;
    document = `<w:document ${NAMESPACES}>${body}</w:document>`;
  }
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
    '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships></pkg:xmlData></pkg:part>" +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    `<pkg:xmlData>${document}</pkg:xmlData></pkg:part>` +
    "</pkg:package>"
  );
}

/** Párrafo con un campo complejo marcado para actualizarse (Word lo recalcula con F9 o al abrir). */
export function fieldParagraph(code: string, placeholder = "Actualice el campo (F9)"): string {
  return (
    '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> ${escapeXml(code.trim())} </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:r><w:t xml:space="preserve">${escapeXml(placeholder)}</w:t></w:r>` +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
  );
}

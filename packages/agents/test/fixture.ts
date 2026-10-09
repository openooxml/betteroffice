import JSZip from 'jszip';

export async function fixture(extraParagraphs = 0): Promise<Uint8Array> {
  const zip = new JSZip();
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${r}/styles" Target="styles.xml"/><Relationship Id="header" Type="${r}/header" Target="header1.xml"/></Relationships>`);
  zip.file('word/styles.xml', `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>`);
  zip.file('word/header1.xml', `<w:hdr xmlns:w="${w}"><w:p><w:r><w:t>Internal — Project Aurora</w:t></w:r></w:p></w:hdr>`);
  zip.file('customXml/preserved.xml', '<payload>Keep these exact bytes.</payload>');
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Executive summary</w:t></w:r></w:p>
    <w:p><w:r><w:t>Project Aurora revenue is </w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="0066AA"/></w:rPr><w:t>€4.2 million</w:t></w:r><w:r><w:t>. Delivery is scheduled for 15 October 2026.</w:t></w:r></w:p>
    <w:p><w:r><w:t>The forecast is draft; the appendix is draft.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After 😀 café İSTANBUL</w:t></w:r></w:p>
    <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Table target: approved</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    ${Array.from({ length: extraParagraphs }, (_, i) => `<w:p><w:r><w:t>Appendix record ${i + 1}: Project Aurora risk assessment remains unchanged.</w:t></w:r></w:p>`).join('')}
    <w:sectPr><w:headerReference w:type="default" r:id="header"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720"/></w:sectPr>
  </w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

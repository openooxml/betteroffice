import JSZip from 'jszip';

export async function hostProposalWorkbook(): Promise<Uint8Array> {
  const zip = new JSZip();
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rels}"><sheets><sheet name="Review" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rels}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetData><row r="3"><c r="B3"><v>100</v></c></row><row r="100"><c r="A100"><v>1</v></c></row></sheetData></worksheet>`);
  zip.file('xl/worksheets/sheet2.xml', `<worksheet xmlns="${main}"><sheetData/></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

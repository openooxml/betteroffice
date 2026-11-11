import JSZip from 'jszip';

export async function fontPreflightFixture(): Promise<Uint8Array> {
  const s = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (entries: [string, string, string][]) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map(([id, kind, target]) => `<Relationship Id="${id}" Type="${r}/${kind}" Target="${target}"/>`).join('')}</Relationships>`;
  const zip = new JSZip();
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + [
      ['/xl/workbook.xml', 'sheet.main'], ['/xl/worksheets/one.xml', 'worksheet'], ['/xl/styles.xml', 'styles'],
    ].map(([part, kind]) => `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${kind}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': rels([['main', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': `<workbook xmlns="${s}" xmlns:r="${r}"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Sheet" sheetId="1" r:id="sheet"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['sheet', 'worksheet', 'worksheets/one.xml'], ['styles', 'styles', 'styles.xml']]),
    'xl/worksheets/one.xml': `<worksheet xmlns="${s}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>漢かな</t></is></c><c r="B1" s="1"><v>12</v></c><c r="U1" s="2"><v>99</v></c></row></sheetData></worksheet>`,
    'xl/styles.xml': `<styleSheet xmlns="${s}"><fonts><font><name val="Base"/></font><font><name val="Styled"/><b/><i/></font><font><name val="Unused"/></font></fonts><cellXfs><xf fontId="0"/><xf fontId="1"/><xf fontId="2"/></cellXfs></styleSheet>`,
  };
  for (const [name, xml] of Object.entries(parts)) zip.file(name, xml);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { sectionPackage } from '../../docx/src/yrs/sectionTestPackage';

const namespaces = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';
const textbox = '<w:p w14:paraId="10000011"><w:r><w:t>Textbox text</w:t></w:r></w:p>';
const drawing = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="1A2B3C4D" wp14:editId="5E6F7A8B"><wp:extent cx="914400" cy="457200"/><wp:docPr id="41" name="Text Box 41"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${textbox}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`;
const alternate = `<w:p w14:paraId="10000010"><w:r><w:t xml:space="preserve">Caption </w:t></w:r><w:r><w:rPr><w:noProof/><w:lang w:eastAsia="en-GB"/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback><w:pict><v:shape style="width:72pt;height:36pt"><v:textbox><w:txbxContent>${textbox}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
const connector = '<w:p w14:paraId="10000020"><w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="2B3C4D5E" wp14:editId="6F7A8B9C"><wp:extent cx="914400" cy="12700"/><wp:docPr id="42" name="Straight Connector 42"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
const table = '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/></w:tcPr><w:p w14:paraId="10000030"><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="eastAsia"/></w:rPr><w:t>Shaded cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

async function fixture(boundary = false) {
  const zip = await JSZip.loadAsync(await sectionPackage(boundary));
  const xml = (await zip.file('word/document.xml')!.async('string'))
    .replace('<w:body>', `<w:body>${alternate}${table}${connector}`)
    .replace('<w:document ', `<w:document ${namespaces} `);
  zip.file('word/document.xml', xml);
  return { bytes: await zip.generateAsync({ type: 'uint8array' }), xml };
}

const paragraph = (xml: string, id: string) => xml.match(new RegExp(`<w:p w14:paraId="${id}">.*?</w:p>`))![0];

test('inserted paragraphs with an added section keep untouched blocks byte for byte', async () => {
  const doc = await openDocx((await fixture()).bytes);
  try {
    const item = (text: string) => doc.list().items.find(entry => entry.text === text)!;
    await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: item('First').ref, paragraphs: [{ text: 'Inserted' }] }] });
    await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_section_break', before: item('Second').ref, kind: 'nextPage' }] });
    await doc.edit({ version: doc.list().version, edits: [{ op: 'set_section', section: 2, orientation: 'landscape', pageWidth: 16838, pageHeight: 11906 }] });
    const bytes = await doc.export();
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    for (const block of [alternate, connector, table]) expect(xml).toContain(block);
    const reopened = await openDocx(bytes);
    try {
      const outline = reopened.list();
      expect(outline.items.map(entry => entry.text)).toEqual(expect.arrayContaining(['First', 'Inserted', 'Second', 'Third']));
      expect(outline.sections.items).toHaveLength(2);
      expect(outline.sections.items[1]).toMatchObject({ number: 2, orientation: 'landscape', pageWidth: 16838, pageHeight: 11906, firstParagraph: outline.items.find(entry => entry.text === 'Second')!.ref });
    } finally { reopened.close(); }
  } finally { doc.close(); }
});

test('inserted paragraphs keep an existing section on its paragraph past textbox paragraphs', async () => {
  const source = await fixture(true);
  const doc = await openDocx(source.bytes);
  try {
    const item = (text: string) => doc.list().items.find(entry => entry.text === text)!;
    await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_paragraphs', after: item('First').ref, paragraphs: [{ text: 'Inserted' }] }] });
    await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_section_break', before: item('Second').ref, kind: 'nextPage' }] });
    const bytes = await doc.export();
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    for (const block of [alternate, connector, table, paragraph(source.xml, '10000001'), paragraph(source.xml, '10000003')]) expect(xml).toContain(block);
    for (const kept of paragraph(source.xml, '10000002').match(/<w:pg(Sz|Mar) [^>]*>/g)!) expect(paragraph(xml, '10000002')).toContain(kept);
    expect(xml.match(/<w:sectPr[ >]/g)).toHaveLength(3);
    const reopened = await openDocx(bytes);
    try {
      const outline = reopened.list();
      const ref = (text: string) => outline.items.find(entry => entry.text === text)!.ref;
      expect(outline.sections.items.map(section => [section.firstParagraph, section.lastParagraph]).slice(1)).toEqual([[ref('Second'), ref('Second')], [ref('Third'), ref('Third')]]);
      expect(outline.sections.items[0].lastParagraph).toBe(ref('Inserted'));
    } finally { reopened.close(); }
  } finally { doc.close(); }
});

test('a section added beside an edited textbox paragraph lands on its own paragraph', async () => {
  const source = await fixture();
  const doc = await openDocx(source.bytes);
  try {
    const item = (text: string) => doc.list().items.find(entry => entry.text === text)!;
    await doc.edit({ version: doc.list().version, edits: [{ op: 'format_text', match: doc.grep({ query: 'Caption' }).matches[0].match, bold: true }] });
    await doc.edit({ version: doc.list().version, edits: [{ op: 'insert_section_break', before: item('Third').ref, kind: 'nextPage' }] });
    const bytes = await doc.export();
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    for (const block of [connector, table, paragraph(source.xml, '10000001'), paragraph(source.xml, '10000003')]) expect(xml).toContain(block);
    const reopened = await openDocx(bytes);
    try {
      const outline = reopened.list();
      const ref = (text: string) => outline.items.find(entry => entry.text === text)!.ref;
      expect(outline.sections.items.map(section => [section.lastParagraph, section.number])).toEqual([[ref('Second'), 1], [ref('Third'), 2]]);
    } finally { reopened.close(); }
  } finally { doc.close(); }
});

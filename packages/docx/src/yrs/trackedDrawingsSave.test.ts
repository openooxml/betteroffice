import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { assertDrawingInvariant } from '../docx/__testSupport__/drawingInvariant';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx } from './index';
import { yrsToDocument } from './yrsToDocument';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"';
const drawing = (text: string) => '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="20" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvPr id="20"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr>'
  + `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`;
const alternate = (text: string) => `<mc:AlternateContent><mc:Choice Requires="wps">${drawing(text)}</mc:Choice><mc:Fallback><w:pict><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent>`;
const ins = '<w:ins w:id="2" w:author="Editor" w:date="2018-07-05T14:06:00Z">';
const del = '<w:del w:id="3" w:author="Editor" w:date="2018-07-05T14:06:00Z">';

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  add('word/document.xml', `<w:document ${NS}><w:body><w:p w14:paraId="00000001"><w:r><w:t>Before</w:t></w:r></w:p>`
    + `<w:p w14:paraId="00000002"><w:r><w:t>Host</w:t></w:r>${ins}<w:r>${alternate('Inserted box')}</w:r><w:r>${drawing('Inserted drawing')}</w:r></w:ins>`
    + `${del}<w:r>${alternate('Deleted box')}</w:r></w:del></w:p><w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function assertKept(source: Uint8Array, saved: Uint8Array): void {
  assertDrawingInvariant(source, saved);
  const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
  for (const text of ['Inserted box', 'Inserted drawing', 'Deleted box']) expect(xml, xml).toContain(text);
  const inserted = xml.indexOf('<w:ins ');
  expect(inserted).toBeGreaterThanOrEqual(0);
  expect(xml.indexOf('Inserted box')).toBeGreaterThan(inserted);
  expect(xml.indexOf('Inserted box')).toBeLessThan(xml.indexOf('</w:ins>', inserted));
  const deleted = xml.indexOf('<w:del ');
  expect(deleted).toBeGreaterThanOrEqual(0);
  expect(xml.indexOf('Deleted box', deleted)).toBeLessThan(xml.indexOf('</w:del>', deleted));
}

it('parse and repack keep text boxes inside tracked changes', async () => {
  const source = fixture();
  assertKept(source, new Uint8Array(await repackDocx(await parseDocx(source.slice().buffer, { preloadFonts: false }))));
});

for (const seeder of ['native', 'projected'] as const) {
  for (const edit of [false, true]) {
    it(`${seeder} session keeps text boxes inside tracked changes${edit ? ' after an edit' : ''}`, async () => {
      const source = fixture();
      const base = await parseDocx(source.slice().buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (seeder === 'native') session.openDocx(source, true);
        else documentToYrs(session, base);
        if (edit) {
          const before = session.paragraphs('body').find(entry => entry.text === 'Before')!;
          session.insertText({ story: 'body', paraId: before.paraId, offset: 0 }, 'x ');
        }
        const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
        if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
        for (const saved of outputs) assertKept(source, saved);
      } finally { session.destroy(); }
    });
  }
}

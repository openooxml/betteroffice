import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
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
const shape = (id: number, inner: string) => `<wps:wsp><wps:cNvPr id="${id}" name="Shape ${id}"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr>${inner}<wps:bodyPr/></wps:wsp>`;
const textBox = '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="20" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
  + shape(21, '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx>') + '</a:graphicData></a:graphic></wp:inline></w:drawing>';
const fallback = '<mc:Fallback><w:pict><v:rect style="width:144pt;height:72pt"/></w:pict></mc:Fallback>';

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
    + `<w:p w14:paraId="00000002"><w:bookmarkStart w:id="0" w:name="Lead"/><w:r><w:t>Host</w:t></w:r><w:bookmarkEnd w:id="0"/>`
    + '<w:r><w:br/></w:r>'
    + `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBox}</mc:Choice>${fallback}</mc:AlternateContent></w:r></w:p><w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function assertKept(saved: Uint8Array): void {
  const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
  expect(xml.split('>Boxed<').length - 1, xml).toBe(1);
  expect(xml.indexOf('<w:br/>'), xml).toBeLessThan(xml.indexOf('>Boxed<'));
}

it('parse and repack keep a text box that follows bookmarks', async () => {
  assertKept(new Uint8Array(await repackDocx(await parseDocx(fixture().slice().buffer, { preloadFonts: false }))));
});

for (const seeder of ['native', 'projected'] as const) {
  it(`${seeder} session keeps a text box that follows bookmarks when its paragraph is edited`, async () => {
    const source = fixture();
    const base = await parseDocx(source.slice().buffer, { preloadFonts: false });
    const session = await createYrsSession();
    try {
      if (seeder === 'native') session.openDocx(source, true);
      else documentToYrs(session, base);
      const host = session.paragraphs('body').find(entry => entry.text.startsWith('Host'))!;
      session.insertText({ story: 'body', paraId: host.paraId, offset: 0 }, 'x ');
      const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
      if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
      for (const saved of outputs) assertKept(saved);
    } finally { session.destroy(); }
  });
}

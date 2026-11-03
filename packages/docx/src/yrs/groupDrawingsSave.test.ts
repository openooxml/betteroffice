import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"';
const GROUP = 'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
const shape = (id: number, inner: string) => `<wps:wsp><wps:cNvPr id="${id}" name="Shape ${id}"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr>${inner}<wps:bodyPr/></wps:wsp>`;
const group = '<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="30" name="Group"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>'
  + shape(31, '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Grouped</w:t></w:r></w:p></w:txbxContent></wps:txbx>')
  + shape(32, '').replace('<a:prstGeom prst="rect"/></wps:spPr>', '<a:prstGeom prst="rect"/><a:blipFill><a:blip r:embed="img"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></wps:spPr>')
  + '</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>';
const textBox = '<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="20" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
  + shape(21, '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx>') + '</a:graphicData></a:graphic></wp:inline></w:drawing>';
const fallback = '<mc:Fallback><w:pict><v:rect style="width:144pt;height:72pt"/></w:pict></mc:Fallback>';

function fixture(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  add('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="img" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pixel.png"/></Relationships>');
  parts.set('word/media/pixel.png', Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='), c => c.charCodeAt(0)));
  add('word/document.xml', `<w:document ${NS}><w:body><w:p w14:paraId="00000001"><w:r><w:t>Before</w:t></w:r></w:p>`
    + `<w:p w14:paraId="00000002"><w:bookmarkStart w:id="0" w:name="Lead"/><w:r><w:t>Host</w:t></w:r><w:bookmarkEnd w:id="0"/>`
    + `<w:r><mc:AlternateContent><mc:Choice Requires="wpg">${group}</mc:Choice>${fallback}</mc:AlternateContent></w:r>`
    + `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${textBox}</mc:Choice>${fallback}</mc:AlternateContent></w:r></w:p><w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function assertKept(saved: Uint8Array): void {
  const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
  expect(xml.split(GROUP).length - 1, xml).toBe(1);
  expect(xml.split('<wps:wsp>').length - 1, xml).toBe(3);
  expect(xml.split('r:embed="img"').length - 1, xml).toBe(1);
  for (const text of ['>Grouped<', '>Boxed<']) expect(xml.split(text).length - 1, xml).toBe(1);
}

it('parse and repack keep a group and a text box', async () => {
  assertKept(new Uint8Array(await repackDocx(await parseDocx(fixture().slice().buffer, { preloadFonts: false }))));
});

for (const seeder of ['native', 'projected'] as const) {
  it(`${seeder} session repack keeps a group and a text box when their paragraph is edited`, async () => {
    const source = fixture();
    const base = await parseDocx(source.slice().buffer, { preloadFonts: false });
    const session = await createYrsSession();
    try {
      if (seeder === 'native') session.openDocx(source, true);
      else documentToYrs(session, base);
      const host = session.paragraphs('body').find(entry => entry.text.startsWith('Host'))!;
      session.insertText({ story: 'body', paraId: host.paraId, offset: 0 }, 'x ');
      assertKept(new Uint8Array(await repackDocx(yrsToDocument(session, base))));
    } finally { session.destroy(); }
  });
}

const scaled = `<w:drawing><wp:inline><wp:extent cx="5486400" cy="5486400"/><wp:docPr id="40" name="Scaled"/><a:graphic><a:graphicData ${GROUP}><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="5486400"/><a:chOff x="105048" y="105048"/><a:chExt cx="2750159" cy="2750159"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:cNvPr id="41" name="Part"/><wps:spPr><a:xfrm><a:off x="105048" y="105048"/><a:ext cx="50000" cy="50000"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>`;

it('a whole-part save keeps a group whose child sits on the scaled child origin', async () => {
  const parts = new Map<string, Uint8Array>(Object.entries(unzipContainer(fixture())));
  parts.set('word/document.xml', toBytes(`<w:document ${NS}><w:body><w:p w14:paraId="00000002"><w:r><w:t>Host</w:t></w:r><w:r>${scaled}</w:r></w:p><w:sectPr/></w:body></w:document>`));
  const source = new Uint8Array(rezipPartsToArrayBuffer(parts));
  const document = await parseDocx(source.slice().buffer, { preloadFonts: false });
  const body = document.package.document;
  body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
  const saved = new Uint8Array(await repackDocx(document));
  const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
  expect(xml).toContain('w:top="2000"');
  expect(xml).toContain(scaled);
});

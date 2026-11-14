import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '.';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { repackDocx } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';

beforeAll(async () => {
  await preloadParseWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/parse/docx_parse_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
const GROUP = 'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
const scaled = `<w:drawing><wp:inline><wp:extent cx="5486400" cy="5486400"/><wp:docPr id="40" name="Scaled"/><a:graphic><a:graphicData ${GROUP}><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="5486400"/><a:chOff x="105048" y="105048"/><a:chExt cx="2750159" cy="2750159"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:cNvPr id="41" name="Part"/><wps:spPr><a:xfrm><a:off x="105048" y="105048"/><a:ext cx="50000" cy="50000"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>`;

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
  add('word/document.xml', `<w:document ${NS}><w:body><w:p><w:r><w:t>Host</w:t></w:r><w:r>${scaled}</w:r></w:p><w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

it('the standalone parser and writer keep a group whose child sits on the scaled child origin', async () => {
  const document = await parseDocx(fixture().slice().buffer, { preloadFonts: false });
  const body = document.package.document;
  body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
  const saved = new Uint8Array(await repackDocx(document));
  const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
  expect(xml).toContain('w:top="2000"');
  expect(xml).toContain(scaled);
});

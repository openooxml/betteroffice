import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createYrsSession, saveYrsDocx } from './index';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function documentBytes(attributeValue: string): Uint8Array<ArrayBuffer> {
  const extra = '<w:bookmarkStart w:id="9" w:name="A_to_B"/><w:bookmarkEnd w:id="9"/>' +
    '<w:r><w:t xml:space="preserve"> </w:t></w:r><w:proofErr w:type="gramStart"/>' +
    '<w:customXml w:uri="urn:x" w:element="note"><w:customXmlPr>' +
    `<w:attr w:name="rule" w:val="${attributeValue}"/></w:customXmlPr></w:customXml>`;
  const body = '<w:p w14:paraId="0000A001"><w:r><w:rPr><w:noProof/></w:rPr><w:t>First</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tblPrEx><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPrEx>' +
    '<w:tc><w:p w14:paraId="0000A002"><w:r><w:rPr><w:noProof/></w:rPr><w:t>Cell one</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p w14:paraId="0000A003"><w:smartTag w:uri="urn:x" w:element="place"><w:r><w:t>Cell two</w:t></w:r></w:smartTag></w:p></w:tc></w:tr></w:tbl>' +
    `<w:p w14:paraId="0000A004">${extra}<w:r><w:t>Last</w:t></w:r></w:p>` +
    '<w:p w14:paraId="0000A005"><w:r><w:rPr><w:noProof/></w:rPr><w:t>Untouched</w:t></w:r></w:p>';
  const parts: PartsMap = new Map();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:wp="${WP}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function boldCount(bytes: Uint8Array<ArrayBuffer>, clientId: number): Promise<number> {
  const session = await createYrsSession({ clientId });
  try {
    session.seedFromDocx(bytes);
    const paragraphs = session.paragraphs('body');
    const first = paragraphs.find(({ text }) => text.includes('First'))!;
    const last = paragraphs.find(({ text }) => text.includes('Last'))!;
    session.formatRange({
      story: 'body',
      start: { paraId: first.paraId, offset: 0 },
      end: { paraId: last.paraId, offset: 4 },
    }, { bold: true });
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
    return xml.split('<w:b/>').length - 1;
  } finally {
    session.destroy();
  }
}

test('bold across a table is saved when an attribute value contains a literal greater-than sign', async () => {
  const plain = await boldCount(documentBytes('a - b'), 98400);
  expect(plain).toBe(2);
  expect(await boldCount(documentBytes('a -> b'), 98401)).toBe(plain);
}, 30_000);

import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession, type YrsSession } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const paragraph = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const TABLE = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraph('Cell')}</w:tc></w:tr></w:tbl>`;
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const AFTER = 'After table';

const DOCX = new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
  ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraph('Before')}${TABLE}${paragraph(AFTER)}${SECT}</w:body></w:document>`)],
])));

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

const open: YrsSession[] = [];
afterEach(() => { open.splice(0).forEach(session => session.destroy()); });

async function afterTable() {
  const session = await createYrsSession();
  open.push(session);
  session.seedFromDocx(DOCX);
  const { paraId, text } = session.paragraphs('body').find(entry => entry.text === AFTER)!;
  const { length } = session.paragraphSpans('body').find(span => span.paraId === paraId)!;
  return { session, paraId, text, length };
}

function boldMarks(session: YrsSession, paraId: string): string {
  const segments = session.storySegments('body');
  const end = segments.findIndex(segment => segment.kind === 'pilcrow' && segment.paraId === paraId);
  let start = end;
  while (start > 0 && segments[start - 1]!.kind === 'text') start--;
  return segments.slice(start, end)
    .map(segment => segment.kind === 'text' ? (segment.attributes.bold === true ? 'B' : '-').repeat(segment.text.length) : '')
    .join('');
}

test('a paragraph after a table counts the table as one offset unit before its text', async () => {
  const { text, length } = await afterTable();
  expect(text).toBe(AFTER);
  expect(length).toBe(text.length + 1);
});

test('offsets 0 to 1 cover the table and none of the text after it', async () => {
  const { session, paraId } = await afterTable();
  session.formatRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 1 } }, { bold: true });
  expect(boldMarks(session, paraId)).toBe('-'.repeat(AFTER.length));
});

test('formatting from offset 1 to the span length covers all of the text after a table', async () => {
  const { session, paraId, length } = await afterTable();
  session.formatRange({ story: 'body', start: { paraId, offset: 1 }, end: { paraId, offset: length } }, { bold: true });
  expect(boldMarks(session, paraId)).toBe('B'.repeat(AFTER.length));
});

test('formatting from offset 1 to the text length stops before the last character after a table', async () => {
  const { session, paraId, text } = await afterTable();
  session.formatRange({ story: 'body', start: { paraId, offset: 1 }, end: { paraId, offset: text.length } }, { bold: true });
  expect(boldMarks(session, paraId)).toBe(`${'B'.repeat(AFTER.length - 1)}-`);
});

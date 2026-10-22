import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const start = '<w:bookmarkStart w:id="1" w:name="mark"/>';
const end = '<w:bookmarkEnd w:id="1"/>';
const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const single = `<w:p w14:paraId="10000001">${run('ab')}${start}${run('cd')}${end}${run('ef')}</w:p>`;
const spanning = `<w:p w14:paraId="10000001">${run('ab')}${start}${run('cd')}</w:p><w:p w14:paraId="10000002">${run('ef')}${end}${run('gh')}</w:p>`;
const second = `<w:p w14:paraId="10000001">${run('ab')}</w:p><w:p w14:paraId="10000002">${run('c')}${start}${run('d')}${end}${run('e')}</w:p>`;
const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p w14:paraId="20000001">${run('T')}</w:p></w:tc></w:tr></w:tbl>`;
const collapsed = `<w:p w14:paraId="10000001">${run('ab')}${start}${end}${run('cd')}</w:p>`;

function fixture(body: string): Uint8Array<ArrayBuffer> {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`)],
  ])));
}

function marked(bytes: Uint8Array): string {
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  const body = xml.slice(xml.indexOf('<w:body>'), xml.lastIndexOf('<w:sectPr'));
  return body
    .replace(/<w:bookmarkStart\b[^>]*\/>/g, '[')
    .replace(/<w:bookmarkEnd\b[^>]*(?:\/>|><\/w:bookmarkEnd>)/g, ']')
    .replace(/<\/w:p>/g, '|')
    .replace(/<[^>]*>/g, '');
}

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

type Edit = (session: YrsSession, paraIds: string[]) => void;
const insert = (para: number, offset: number, text = 'X'): Edit => (session, paraIds) => {
  session.insertText({ story: 'body', paraId: paraIds[para]!, offset }, text);
};
const remove = (para: number, from: number, to: number, toPara = para): Edit => (session, paraIds) => {
  session.deleteRange({ story: 'body', start: { paraId: paraIds[para]!, offset: from }, end: { paraId: paraIds[toPara]!, offset: to } });
};
const split = (para: number, offset: number): Edit => (session, paraIds) => {
  session.splitParagraph({ story: 'body', paraId: paraIds[para]!, offset });
};
const rejected = (para: number, offset: number): Edit => (session, paraIds) => {
  const { revisionId } = session.insertText({ story: 'body', paraId: paraIds[para]!, offset }, 'X', { name: 'A', date: '2026-01-01T00:00:00Z' });
  session.rejectChange({ revisionId: revisionId! });
};

const cases: Array<{ name: string; body: string; edits: Edit[]; expected: string }> = [
  { name: 'typing before a bookmark', body: single, edits: [insert(0, 0), insert(0, 1)], expected: 'XXab[cd]ef|' },
  { name: 'typing inside a bookmark', body: single, edits: [insert(0, 3), insert(0, 4)], expected: 'ab[cXXd]ef|' },
  { name: 'typing at a bookmark start', body: single, edits: [insert(0, 2)], expected: 'ab[Xcd]ef|' },
  { name: 'typing at a bookmark end', body: single, edits: [insert(0, 4)], expected: 'ab[cd]Xef|' },
  { name: 'typing after a bookmark', body: single, edits: [insert(0, 5)], expected: 'ab[cd]eXf|' },
  { name: 'typing at a collapsed bookmark', body: collapsed, edits: [insert(0, 0), insert(0, 3)], expected: 'Xab[]Xcd|' },
  { name: 'deleting before a bookmark', body: single, edits: [remove(0, 0, 1)], expected: 'b[cd]ef|' },
  { name: 'deleting inside a bookmark', body: single, edits: [remove(0, 2, 3)], expected: 'ab[d]ef|' },
  { name: 'deleting across a bookmark start', body: single, edits: [remove(0, 1, 3)], expected: 'a[d]ef|' },
  { name: 'deleting a whole bookmarked span', body: single, edits: [remove(0, 2, 4)], expected: 'ab[]ef|' },
  { name: 'typing in both paragraphs of a spanning bookmark', body: spanning, edits: [insert(0, 0), insert(0, 4), insert(1, 0), insert(1, 4)], expected: 'Xab[cXd|Xef]gXh|' },
  { name: 'splitting a paragraph inside a bookmark', body: single, edits: [split(0, 3)], expected: 'ab[c|d]ef|' },
  { name: 'splitting a paragraph at a bookmark start', body: single, edits: [split(0, 2)], expected: 'ab|[cd]ef|' },
  { name: 'splitting a paragraph at a bookmark end', body: single, edits: [split(0, 4)], expected: 'ab[cd]|ef|' },
  { name: 'splitting a paragraph before a bookmark', body: single, edits: [split(0, 1)], expected: 'a|b[cd]ef|' },
  { name: 'splitting a paragraph after a bookmark', body: single, edits: [split(0, 5)], expected: 'ab[cd]e|f|' },
  { name: 'merging the paragraphs of a spanning bookmark', body: spanning, edits: [remove(0, 4, 0, 1)], expected: 'ab[cdef]gh|' },
  { name: 'merging into a paragraph before a bookmark', body: second, edits: [remove(0, 2, 0, 1)], expected: 'abc[d]e|' },
  { name: 'rejecting a suggested insertion before a bookmark', body: single, edits: [insert(0, 0), rejected(0, 1)], expected: 'Xab[cd]ef|' },
  { name: 'deleting in both paragraphs of a spanning bookmark', body: spanning, edits: [remove(0, 0, 1), remove(1, 0, 1)], expected: 'b[cd|f]gh|' },
  // Session offsets count the leading table as one unit, so its paragraph's text starts at 1.
  { name: 'typing inside a bookmark after a leading table', body: table + single, edits: [insert(0, 4)], expected: 'T|ab[cXd]ef|' },
  { name: 'deleting inside a bookmark after a leading table', body: table + single, edits: [remove(0, 3, 4)], expected: 'T|ab[d]ef|' },
  { name: 'splitting a paragraph inside a bookmark after a leading table', body: table + single, edits: [split(0, 4)], expected: 'T|ab[c|d]ef|' },
  { name: 'splitting a paragraph at a bookmark start after a leading table', body: table + single, edits: [split(0, 3)], expected: 'T|ab|[cd]ef|' },
  { name: 'merging the paragraphs of a spanning bookmark after a leading table', body: table + spanning, edits: [remove(0, 5, 0, 1)], expected: 'T|ab[cdef]gh|' },
  { name: 'merging a bookmark into a paragraph with a leading table', body: table + second, edits: [remove(0, 3, 0, 1)], expected: 'T|abc[d]e|' },
  { name: 'deleting across paragraphs of a spanning bookmark after a leading table', body: table + spanning, edits: [remove(0, 4, 1, 1)], expected: 'T|ab[cf]gh|' },
];

for (const seeder of ['native', 'projected'] as const) {
  for (const { name, body, edits, expected } of cases) {
    it(`${seeder}: ${name} keeps the bookmark range in the saved document`, async () => {
      const bytes = fixture(body);
      const base = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (seeder === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, base);
        const paraIds = session.paragraphs('body').map(paragraph => paragraph.paraId).filter(paraId => paraId !== '20000001');
        for (const edit of edits) edit(session, paraIds);
        expect(marked(new Uint8Array(await repackDocx(yrsToDocument(session, base))))).toBe(expected);
        if (seeder === 'native') expect(marked((await saveYrsDocx(session)).bytes)).toBe(expected);
      } finally {
        session.destroy();
      }
    });
  }
}

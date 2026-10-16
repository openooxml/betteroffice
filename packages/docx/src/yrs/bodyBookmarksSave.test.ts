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

const tableStart = `<w:bookmarkStart w:name='Table&amp;Range' w:colLast="2" w:id = '7' w:colFirst="0" w:displacedByCustomXml="next" x:flag="a&#38;b" />`;
const tableEnd = `<w:bookmarkEnd w:displacedByCustomXml='prev' w:id="7"></w:bookmarkEnd>`;
const sectionStart = '<w:bookmarkStart w:id="8" w:name="SectionRange"/>';
const sectionEnd = '<w:bookmarkEnd w:id="8"/>';
const paragraph = '<w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>';
const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>';
const section = '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>';

function fixture(): Uint8Array<ArrayBuffer> {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:bookmark"><w:body>${paragraph}${tableStart}${table}${tableEnd}${sectionStart}${section}${sectionEnd}<w:sectPr/></w:body></w:document>`)],
  ])));
}

function assertBookmarks(bytes: Uint8Array): void {
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  let offset = 0;
  for (const fragment of ['Edited', tableStart, '<w:tbl>', '</w:tbl>', tableEnd, sectionStart, '>Section<', sectionEnd]) {
    expect(xml.split(fragment)).toHaveLength(2);
    const index = xml.indexOf(fragment, offset);
    expect(index).toBeGreaterThanOrEqual(offset);
    offset = index + fragment.length;
  }
}

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

for (const seeder of ['native', 'projected']) {
  it(`${seeder} keeps body bookmarks in place after repeated text edits and saves`, async () => {
    const bytes = fixture();
    const base = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession();
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, base);
      const first = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      for (let count = 0; count < 2; count += 1) {
        assertBookmarks(new Uint8Array(await repackDocx(yrsToDocument(session, base))));
        if (seeder === 'native') assertBookmarks((await saveYrsDocx(session)).bytes);
        session.insertText({ story: 'body', paraId: first.paraId, offset: 7 }, '!');
      }
    } finally {
      session.destroy();
    }
  });
}

it('keeps body bookmarks through a table structure rewrite', async () => {
  const session = await createYrsSession();
  try {
    session.seedFromDocx(fixture());
    const first = session.paragraphs('body')[0]!;
    session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
    session.insertColumn({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'right');
    assertBookmarks((await saveYrsDocx(session)).bytes);
  } finally {
    session.destroy();
  }
});

for (const seeder of ['native', 'projected']) {
  it(`${seeder} refuses a save after deleting the table between bookmark boundaries`, async () => {
    const bytes = fixture();
    const base = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession();
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, base);
      session.deleteTable({ story: 'body', tableIndex: 0 });
      expect(() => yrsToDocument(session, base)).toThrow('bookmark boundaries');
      if (seeder === 'native') await expect(saveYrsDocx(session)).rejects.toThrow('bookmark boundaries');
    } finally {
      session.destroy();
    }
  });
}

it('keeps bookmark positions when an unrelated paragraph is inserted and deleted', async () => {
  const session = await createYrsSession();
  try {
    session.seedFromDocx(fixture());
    const first = session.paragraphs('body')[0]!;
    session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
    expect(session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'insertParagraphs', target: { story: 'body', paraId: first.paraId }, at: 'end', paragraphs: [{ text: 'Inserted' }] }],
    }).ok).toBe(true);
    assertBookmarks((await saveYrsDocx(session)).bytes);
    const inserted = session.paragraphs('body').find(({ text }) => text === 'Inserted')!;
    expect(session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'deleteParagraphs', story: 'body', firstParaId: inserted.paraId, lastParaId: inserted.paraId }],
    }).ok).toBe(true);
    assertBookmarks((await saveYrsDocx(session)).bytes);
  } finally {
    session.destroy();
  }
});

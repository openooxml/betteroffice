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

const tableStart = `<w:bookmarkStart w:name='Table&amp;Range' w:colLast="2" x:flag="w:id='99' a&#38;b" w:id = '&#55;' w:colFirst="0" w:displacedByCustomXml="next" />`;
const tableEnd = `<w:bookmarkEnd w:displacedByCustomXml='prev' w:id="7"></w:bookmarkEnd>`;
const sectionStart = '<w:bookmarkStart w:id="8" w:name="SectionRange"/>';
const sectionEnd = '<w:bookmarkEnd w:id="8"/>';
const paragraph = '<w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>';
const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>';
const section = '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>';

function fixture(body = `${paragraph}${tableStart}${table}${tableEnd}${sectionStart}${section}${sectionEnd}`): Uint8Array<ArrayBuffer> {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:bookmark" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`)],
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

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function assertSavedPair(bytes: Uint8Array, order: string[]): Promise<void> {
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  expect(xml.match(/<w:bookmarkStart\b/g)).toHaveLength(1);
  expect(xml.match(/<w:bookmarkEnd\b/g)).toHaveLength(1);
  expect(xml).toContain(tableStart);
  expect(xml).toContain(tableEnd);
  let offset = 0;
  for (const fragment of order) {
    const position = xml.indexOf(fragment, offset);
    expect(position).toBeGreaterThanOrEqual(offset);
    offset = position + fragment.length;
  }
  expect(xml.indexOf(tableEnd)).toBeLessThan(xml.lastIndexOf('<w:sectPr'));
  const reopened = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
  const raws = reopened.package.document.content.filter(block => block.type === 'rawXml');
  expect(raws.map(block => block.xml)).toEqual([tableStart, tableEnd]);
}

const deletionBody = `${p('Before')}${tableStart}${p('Anchor')}${tableEnd}${p('After')}`;
const cases: Array<{ name: string; body: string; edit: (session: YrsSession) => void; order: string[] }> = [
  {
    name: 'deletes an anchored paragraph', body: deletionBody,
    edit(session) {
      const anchor = session.paragraphs('body').find(para => para.text === 'Anchor')!;
      const span = session.locateParagraph('body', anchor.paraId);
      session.applyRawOps('body', [{ op: 'delete', index: span.start, len: span.end - span.start + 1 }]);
      expect(session.paragraphs('body').some(para => para.paraId === anchor.paraId)).toBe(false);
    },
    order: ['>Before<', tableStart, tableEnd, '>After<'],
  },
  {
    name: 'merges the anchored paragraph into its neighbour', body: deletionBody,
    edit(session) { session.mergeParagraphs('body', session.paragraphs('body')[0]!.paraId); },
    order: ['Anchor', tableStart, tableEnd, '>After<'],
  },
  {
    name: 'deletes a table spanned by a pair', body: `${p('Before')}${tableStart}${table}${tableEnd}${p('After')}`,
    edit(session) { session.deleteTable({ story: 'body', tableIndex: 0 }); },
    order: ['>Before<', tableStart, tableEnd, '>After<'],
  },
  {
    name: 'moves a block across a boundary and collapses an inverted pair', body: deletionBody,
    edit(session) {
      const anchor = session.paragraphs('body').find(para => para.text === 'Anchor')!;
      const segment = session.storySegments('body').find(segment => segment.kind === 'pilcrow' && segment.paraId === anchor.paraId)!;
      if (segment.kind !== 'pilcrow') throw new Error('missing anchor');
      const span = session.locateParagraph('body', anchor.paraId);
      session.applySeedRawOps('body', [
        { op: 'delete', index: span.start, len: span.end - span.start + 1 },
        { op: 'insert', index: 13, text: 'Anchor' },
        { op: 'insertEmbed', index: 19, kind: 'pilcrow', payload: { ...segment.properties, paraId: anchor.paraId } },
      ]);
      expect(session.paragraphs('body').map(para => para.text)).toEqual(['Before', 'After', 'Anchor']);
      expect(session.paragraphs('body')[2]!.paraId).toBe(anchor.paraId);
    },
    order: ['>Before<', '>After<', tableStart, tableEnd, '>Anchor<'],
  },
  {
    name: 'falls back to the previous block at the body tail', body: `${p('Before')}${tableStart}${p('Anchor')}${tableEnd}`,
    edit(session) {
      const anchor = session.paragraphs('body')[1]!;
      session.applyRawOps('body', [{ op: 'delete', index: 7, len: anchor.text.length + 1 }]);
    },
    order: ['>Before<', tableStart, tableEnd],
  },
  {
    name: 'collapses at the body start when no source block survives', body: `${tableStart}${table}${tableEnd}`,
    edit(session) { session.deleteTable({ story: 'body', tableIndex: 0 }); },
    order: [tableStart, tableEnd, '<w:sectPr'],
  },
  {
    name: 'deletes an anchored section paragraph', body: `${p('Before')}${tableStart}${section}${tableEnd}${p('After')}`,
    edit(session) {
      const anchor = session.paragraphs('body').find(para => para.text === 'Section')!;
      const span = session.locateParagraph('body', anchor.paraId);
      session.applyRawOps('body', [{ op: 'delete', index: span.start, len: span.end - span.start + 1 }]);
    },
    order: ['>Before<', tableStart, tableEnd, '>After<'],
  },
];

for (const seeder of ['native', 'projected']) {
  for (const scenario of cases) {
    it(`${seeder} keeps body bookmarks when it ${scenario.name}`, async () => {
      const bytes = fixture(scenario.body);
      const base = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (seeder === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, base);
        scenario.edit(session);
        for (let count = 0; count < 2; count += 1) {
          await assertSavedPair(new Uint8Array(await repackDocx(yrsToDocument(session, base))), scenario.order);
          if (seeder === 'native') await assertSavedPair((await saveYrsDocx(session)).bytes, scenario.order);
        }
      } finally { session.destroy(); }
    });
  }

  it(`${seeder} keeps anchors with repeated paragraph ids, tables, controls and a split`, async () => {
    const repeated = (text: string) => `<w:p w14:paraId="12345678"><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const body = `${repeated('Before')}${tableStart}${repeated('Anchor')}${table}${tableEnd}<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>${p('Control')}</w:sdtContent></w:sdt>${p('After')}`;
    const bytes = fixture(body);
    const base = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession();
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, base);
      const anchor = session.paragraphs('body').find(para => para.text === 'Anchor')!;
      session.splitParagraph({ story: 'body', paraId: anchor.paraId, offset: 3 });
      await assertSavedPair(new Uint8Array(await repackDocx(yrsToDocument(session, base))), ['>Before<', tableStart, '>Anc<', '>hor<', '<w:tbl>', tableEnd, '>Control<']);
      if (seeder === 'native') await assertSavedPair((await saveYrsDocx(session)).bytes, ['>Before<', tableStart, '>Anc<', '>hor<', '<w:tbl>', tableEnd, '>Control<']);
    } finally { session.destroy(); }
  });

  for (const kinds of [['start', 'end'], ['start'], ['end']]) {
    it(`${seeder} restores only missing boundaries when paragraph bookmarks already carry ${kinds.join('/')}`, async () => {
      const bytes = fixture(deletionBody);
      const base = await parseDocx(bytes.buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (seeder === 'native') session.seedFromDocx(bytes);
        else documentToYrs(session, base);
        const anchor = session.paragraphs('body').find(para => para.text === 'Anchor')!;
        const span = session.locateParagraph('body', anchor.paraId);
        session.applyRawOps('body', [{ op: 'setEmbedAttr', index: span.end, key: 'bookmarks', value: kinds.map(kind => ({ id: 7, kind, name: 'Table&Range', offset: kind === 'start' ? 0 : 6 })) }]);
        const outputs: Uint8Array[] = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
        if (seeder === 'native') outputs.push((await saveYrsDocx(session)).bytes);
        for (const saved of outputs) {
          const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
          expect(xml.match(/<w:bookmarkStart\b/g)).toHaveLength(1);
          expect(xml.match(/<w:bookmarkEnd\b/g)).toHaveLength(1);
          expect(xml.indexOf('<w:bookmarkStart')).toBeLessThan(xml.indexOf('<w:bookmarkEnd'));
          expect(xml.indexOf('<w:bookmarkEnd')).toBeLessThan(xml.lastIndexOf('<w:sectPr'));
          await parseDocx(saved.slice().buffer, { preloadFonts: false });
        }
      } finally { session.destroy(); }
    });
  }
}

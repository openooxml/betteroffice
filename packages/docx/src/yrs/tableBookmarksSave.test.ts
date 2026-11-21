import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { assertBookmarkInvariant, collectBookmarkBoundaries } from '../docx/__testSupport__/bookmarkInvariant';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const start = (id: number) => `<w:bookmarkStart w:id="${id}" w:name="Mark${id}"/>`;
const end = (id: number) => `<w:bookmarkEnd w:id="${id}"/>`;
const paragraph = (id: string, text: string) => `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (id: string, text: string) => `<w:tc>${paragraph(id, text)}</w:tc>`;

const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`
  + `${start(1)}<w:tr>${start(3)}${cell('00000011', 'A')}${end(3)}${cell('00000012', 'B')}</w:tr>${start(2)}`
  + `<w:tr>${cell('00000021', 'C')}${start(4)}${cell('00000022', 'D')}${end(4)}</w:tr>${end(2)}`
  + `<w:tr>${cell('00000031', 'E')}${cell('00000032', 'F')}</w:tr>${end(1)}</w:tbl>`;

function fixture(body = table): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  add('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraph('00000001', 'Before')}${body}`
    + `${paragraph('00000002', 'After')}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function documentXml(bytes: Uint8Array): string {
  return new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
}

function assertOrder(xml: string, fragments: string[]): void {
  let offset = 0;
  for (const fragment of fragments) {
    const at = xml.indexOf(fragment, offset);
    expect(at, `${fragment} in ${xml}`).toBeGreaterThanOrEqual(0);
    offset = at + fragment.length;
  }
}

function assertPairs(xml: string): void {
  for (const id of [1, 2, 3, 4]) {
    expect(xml.split(start(id))).toHaveLength(2);
    expect(xml.split(end(id))).toHaveLength(2);
    expect(xml.indexOf(start(id))).toBeLessThan(xml.indexOf(end(id)));
  }
}

const sourceOrder = [
  '<w:tbl>', '</w:tblGrid>', start(1), '<w:tr>', start(3), '>A<', end(3), '>B<', '</w:tr>', start(2),
  '>C<', start(4), '>D<', end(4), '</w:tr>', end(2), '>E<', '>F<', '</w:tr>', end(1), '</w:tbl>', '>After<',
];

it('parse and repack keep bookmarks between rows and cells in place', async () => {
  const source = fixture();
  const saved = new Uint8Array(await repackDocx(await parseDocx(source.slice().buffer, { preloadFonts: false })));
  assertBookmarkInvariant(source, saved);
  assertOrder(documentXml(saved), sourceOrder);
});

type Edit = 'untouched' | 'text' | 'cellText' | 'deleteFirstRow' | 'deleteMiddleRow' | 'insertRow' | 'deleteColumn';

function edit(session: YrsSession, kind: Edit): void {
  const loc = (row: number, column: number) => ({ story: 'body', tableIndex: 0, row, column });
  if (kind === 'text') {
    const after = session.paragraphs('body').find(entry => entry.text === 'After')!;
    session.insertText({ story: 'body', paraId: after.paraId, offset: 0 }, 'x ');
  } else if (kind === 'cellText') {
    const story = session.storyIds().find(id => session.paragraphs(id).some(entry => entry.text === 'C'))!;
    const target = session.paragraphs(story).find(entry => entry.text === 'C')!;
    session.insertText({ story, paraId: target.paraId, offset: 0 }, 'x ');
  } else if (kind === 'deleteFirstRow') {
    session.deleteRow({ anchor: loc(0, 0), head: loc(0, 0) });
  } else if (kind === 'deleteMiddleRow') {
    session.deleteRow({ anchor: loc(1, 0), head: loc(1, 0) });
  } else if (kind === 'insertRow') {
    session.insertRow(loc(0, 0), 'above');
  } else if (kind === 'deleteColumn') {
    session.deleteColumn({ anchor: loc(0, 0), head: loc(0, 0) });
  }
}

const edits: Edit[] = ['untouched', 'text', 'cellText', 'deleteFirstRow', 'deleteMiddleRow', 'insertRow', 'deleteColumn'];

for (const seeder of ['native', 'projected'] as const) {
  for (const kind of edits) {
    it(`${seeder} keeps table and row bookmarks after ${kind}`, async () => {
      const source = fixture();
      const base = await parseDocx(source.slice().buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (seeder === 'native') session.openDocx(source, true);
        else documentToYrs(session, base);
        edit(session, kind);
        const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
        if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
        for (const saved of outputs) {
          assertBookmarkInvariant(source, saved);
          expect(collectBookmarkBoundaries(saved).get('word/document.xml')).toEqual(
            collectBookmarkBoundaries(source).get('word/document.xml'));
          const xml = documentXml(saved);
          assertPairs(xml);
          if (kind === 'untouched' || kind === 'text' || kind === 'cellText') assertOrder(xml, sourceOrder.filter(fragment => fragment !== '>C<' && fragment !== '>After<'));
          if (kind === 'deleteMiddleRow') assertOrder(xml, ['>A<', start(2), '>E<']);
          await parseDocx(saved.slice().buffer, { preloadFonts: false });
        }
      } finally { session.destroy(); }
    });
  }
}

const grid = (columns: number) => `<w:tblPr/><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>`;
const merged = (id: string, text: string, vMerge: 'restart' | 'continue') =>
  `<w:tc><w:tcPr><w:vMerge${vMerge === 'restart' ? ' w:val="restart"' : ''}/></w:tcPr>${paragraph(id, text)}</w:tc>`;
const markedCell = (id: string, text: string, mark: number) =>
  `<w:tc><w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r>${end(mark)}</w:p></w:tc>`;

function pairsOf(xml: string, id: number): void {
  const starts = xml.split(start(id)).length - 1;
  const ends = xml.split(end(id)).length - 1;
  expect(starts, xml).toBe(ends);
  expect(starts, xml).toBeLessThanOrEqual(1);
  if (starts === 1) expect(xml.indexOf(start(id))).toBeLessThan(xml.indexOf(end(id)));
}

const regressions: Array<{ name: string; body: string; edit: (session: YrsSession) => void; kept: boolean; ids: number[] }> = [
  {
    name: 'a continuation cell before a bookmarked cell',
    body: `<w:tbl>${grid(2)}<w:tr>${merged('00000041', 'A', 'restart')}${cell('00000042', 'B')}</w:tr>`
      + `<w:tr>${merged('00000043', '', 'continue')}${start(5)}${markedCell('00000044', 'D', 5)}</w:tr></w:tbl>`,
    edit: (session) => {
      const story = session.storyIds().find(id => session.paragraphs(id).some(entry => entry.text === 'D'))!;
      const target = session.paragraphs(story).find(entry => entry.text === 'D')!;
      session.insertText({ story, paraId: target.paraId, offset: 0 }, 'x ');
    },
    kept: true,
    ids: [5],
  },
  {
    name: 'a row of continuation cells',
    body: `<w:tbl>${grid(1)}<w:tr>${merged('00000051', 'A', 'restart')}</w:tr>`
      + `<w:tr>${merged('00000052', '', 'continue')}${start(6)}</w:tr>${end(6)}</w:tbl>`,
    edit: (session) => {
      const story = session.storyIds().find(id => session.paragraphs(id).some(entry => entry.text === 'A'))!;
      const target = session.paragraphs(story).find(entry => entry.text === 'A')!;
      session.insertText({ story, paraId: target.paraId, offset: 0 }, 'x ');
    },
    kept: true,
    ids: [6],
  },
  {
    name: 'a deleted row whose pair reaches into its cell',
    body: `<w:tbl>${grid(1)}<w:tr>${cell('00000061', 'A')}</w:tr><w:tr>${start(7)}${markedCell('00000062', 'B', 7)}</w:tr></w:tbl>`,
    edit: (session) => session.deleteRow({ anchor: { story: 'body', tableIndex: 0, row: 1, column: 0 }, head: { story: 'body', tableIndex: 0, row: 1, column: 0 } }),
    kept: false,
    ids: [7],
  },
  {
    name: 'a deleted row declaring a namespace',
    body: `<w:tbl>${grid(1)}<w:tr>${cell('00000071', 'A')}</w:tr><w:tr xmlns:ns="urn:example">`
      + `<w:bookmarkStart w:id="8" w:name="Mark8" ns:tag="1"/>${end(8)}${cell('00000072', 'B')}</w:tr></w:tbl>`,
    edit: (session) => session.deleteRow({ anchor: { story: 'body', tableIndex: 0, row: 1, column: 0 }, head: { story: 'body', tableIndex: 0, row: 1, column: 0 } }),
    kept: true,
    ids: [],
  },
  {
    name: 'a deleted row wrapping its boundaries',
    body: `<w:tbl>${grid(1)}<w:tr>${cell('00000081', 'A')}</w:tr><w:tr><w:customXml w:element="wrap">${start(9)}${end(9)}`
      + `${cell('00000082', 'B')}</w:customXml></w:tr></w:tbl>`,
    edit: (session) => session.deleteRow({ anchor: { story: 'body', tableIndex: 0, row: 1, column: 0 }, head: { story: 'body', tableIndex: 0, row: 1, column: 0 } }),
    kept: true,
    ids: [9],
  },
  {
    name: 'boundaries around a continuation cell',
    body: `<w:tbl>${grid(2)}<w:tr>${merged('00000091', 'A', 'restart')}${cell('00000092', 'B')}</w:tr>`
      + `<w:tr>${start(10)}${merged('00000093', '', 'continue')}${end(10)}${cell('00000094', 'D')}</w:tr></w:tbl>`,
    edit: (session) => {
      const story = session.storyIds().find(id => session.paragraphs(id).some(entry => entry.text === 'D'))!;
      const target = session.paragraphs(story).find(entry => entry.text === 'D')!;
      session.insertText({ story, paraId: target.paraId, offset: 0 }, 'x ');
    },
    kept: true,
    ids: [10],
  },
  {
    name: 'every original row replaced',
    body: `<w:tbl>${grid(1)}${start(11)}<w:tr>${cell('000000a1', 'A')}</w:tr>${end(11)}</w:tbl>`,
    edit: (session) => {
      session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'below');
      session.deleteRow({ anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 }, head: { story: 'body', tableIndex: 0, row: 0, column: 0 } });
    },
    kept: true,
    ids: [11],
  },
  {
    name: 'merged cells carrying the partner',
    body: `<w:tbl>${grid(2)}<w:tr>${start(12)}${cell('000000b1', 'A')}${markedCell('000000b2', 'B', 12)}</w:tr></w:tbl>`,
    edit: (session) => session.mergeCells({ anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 }, head: { story: 'body', tableIndex: 0, row: 0, column: 1 } }),
    kept: true,
    ids: [12],
  },
  {
    name: 'a merged-away cell holding both ends',
    body: `<w:tbl>${grid(2)}<w:tr>${cell('000000c1', 'A')}${start(13)}${markedCell('000000c2', 'B', 13)}</w:tr></w:tbl>`,
    edit: (session) => session.mergeCells({ anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 }, head: { story: 'body', tableIndex: 0, row: 0, column: 1 } }),
    kept: true,
    ids: [13],
  },
  {
    name: 'a row inserted inside a vertical merge',
    body: `<w:tbl>${grid(2)}<w:tr>${merged('000000d1', 'A', 'restart')}${cell('000000d2', 'B')}</w:tr>`
      + `<w:tr>${merged('000000d3', '', 'continue')}${start(14)}${markedCell('000000d4', 'D', 14)}</w:tr></w:tbl>`,
    edit: (session) => session.insertRow({ story: 'body', tableIndex: 0, row: 1, column: 1 }, 'above'),
    kept: true,
    ids: [14],
  },
  {
    name: 'a row deleted inside a vertical merge',
    body: `<w:tbl>${grid(2)}<w:tr>${merged('000000e1', 'A', 'restart')}${cell('000000e2', 'B')}</w:tr>`
      + `<w:tr>${merged('000000e3', '', 'continue')}${cell('000000e4', 'C')}</w:tr>`
      + `<w:tr>${merged('000000e5', '', 'continue')}${start(15)}${markedCell('000000e6', 'E', 15)}</w:tr></w:tbl>`,
    edit: (session) => session.deleteRow({ anchor: { story: 'body', tableIndex: 0, row: 1, column: 1 }, head: { story: 'body', tableIndex: 0, row: 1, column: 1 } }),
    kept: true,
    ids: [15],
  },
];

for (const seeder of ['native', 'projected'] as const) {
  for (const regression of regressions) {
    it(`${seeder} keeps valid table bookmark pairs with ${regression.name}`, async () => {
      const source = fixture(regression.body);
      const base = await parseDocx(source.slice().buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (seeder === 'native') session.openDocx(source, true);
        else documentToYrs(session, base);
        regression.edit(session);
        const outputs = [new Uint8Array(await repackDocx(yrsToDocument(session, base)))];
        if (seeder === 'native') outputs.push(new Uint8Array((await saveYrsDocx(session)).bytes));
        for (const saved of outputs) {
          const xml = documentXml(saved);
          for (const id of regression.ids) {
            pairsOf(xml, id);
            if (regression.kept) expect(xml, xml).toContain(start(id));
          }
          if (regression.name.includes('namespace') && xml.includes('ns:tag')) {
            expect(xml).toMatch(/<w:bookmarkStart[^>]*xmlns:ns="urn:example"|<w:(?:document|tbl|tr)[^>]*xmlns:ns="urn:example"/);
          }
          await parseDocx(saved.slice().buffer, { preloadFonts: false });
        }
      } finally { session.destroy(); }
    });
  }
}

const threeRows = `<w:tbl>${grid(1)}<w:tr>${cell('000000f1', 'RowA')}</w:tr>`
  + `<w:tr>${cell('000000f2', 'RowB')}</w:tr>${start(16)}`
  + `<w:tr>${start(17)}<w:tc>${start(18)}${paragraph('000000f3', 'RowC')}${end(18)}</w:tc>${end(17)}</w:tr>${end(16)}</w:tbl>`;

for (const seeder of ['native', 'projected'] as const) {
  for (const first of ['delete', 'insert'] as const) {
    it(`${seeder} keeps row C's bookmarks on row C across two saves after a row ${first}`, async () => {
      const source = fixture(threeRows);
      const base = await parseDocx(source.slice().buffer, { preloadFonts: false });
      const session = await createYrsSession();
      try {
        if (seeder === 'native') session.openDocx(source, true);
        else documentToYrs(session, base);
        const loc = (row: number) => ({ story: 'body', tableIndex: 0, row, column: 0 });
        if (first === 'delete') session.deleteRow({ anchor: loc(0), head: loc(0) });
        else session.insertRow(loc(0), 'below');
        const saved = yrsToDocument(session, base);
        const story = session.storyIds().find(id => session.paragraphs(id).some(entry => entry.text === 'RowB'))!;
        const target = session.paragraphs(story).find(entry => entry.text === 'RowB')!;
        session.insertText({ story, paraId: target.paraId, offset: 0 }, 'x ');
        const bytes = new Uint8Array(await repackDocx(yrsToDocument(session, saved)));
        const xml = documentXml(bytes);
        for (const id of [16, 17, 18]) pairsOf(xml, id);
        assertOrder(xml, ['RowB<', start(16), start(17), start(18), '>RowC<', end(18), end(17), end(16)]);
        expect(xml.indexOf(start(16))).toBeGreaterThan(xml.indexOf('RowB<'));
      } finally { session.destroy(); }
    });
  }
}

import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsCellLoc, type YrsSession, type YrsTableLoc } from './index';
import { sessionUnavailableTablesFrom } from './sessionInternals';
import { savedRootStories, unavailableTablesByProjection } from './tableUnavailable';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const table = (...cells: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${cells.map(() => '<w:gridCol w:w="4680"/>').join('')}</w:tblGrid><w:tr>${cells.join('')}</w:tr></w:tbl>`;
const plain = (name: string) => table(cell(paragraph(`${name} one`)), cell(paragraph(`${name} two`)));
const control = `<w:sdt><w:sdtPr><w:alias w:val="Box"/></w:sdtPr><w:sdtContent>${plain('Control')}${paragraph('In control')}</w:sdtContent></w:sdt>`;

/**
 * Every container a table can sit in: the body, a block content control in a cell, a table in a
 * cell, a header with relationship id `top`, a footnote and a comment.
 */
function containers(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE_DOC}.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="${OFFICE_DOC}.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE_DOC}.wordprocessingml.comments+xml"/></Types>`));
  parts.set('_rels/.rels', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`));
  parts.set('word/_rels/document.xml.rels', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="top" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdF1" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rIdC1" Type="${REL}/comments" Target="comments.xml"/></Relationships>`));
  parts.set('word/header1.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${W}>${plain('Header')}${paragraph('Header text')}</w:hdr>`));
  parts.set('word/footnotes.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${plain('Note')}${paragraph('Note text')}</w:footnote></w:footnotes>`));
  parts.set('word/comments.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${W}><w:comment w:id="0" w:author="Ada" w:date="2026-11-10T00:00:00Z">${plain('Comment')}${paragraph('Comment text')}</w:comment></w:comments>`));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Before</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>${table(cell(`${control}${paragraph('In cell')}`), cell(`${plain('Inner')}${paragraph('Beside inner')}`))}${paragraph('After')}<w:sectPr><w:headerReference w:type="default" r:id="top"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const owned: YrsSession[] = [];

async function session(clientId: number): Promise<YrsSession> {
  const created = await createYrsSession({ clientId });
  owned.push(created);
  return created;
}

function exchange(a: YrsSession, b: YrsSession): void {
  const fromA = a.encodeState();
  const fromB = b.encodeState();
  b.applyUpdate(fromA);
  a.applyUpdate(fromB);
}

function crossing(a: YrsSession, b: YrsSession, story: string, properties = false): void {
  const at: YrsCellLoc = { story, tableIndex: 0, row: 0, column: 0 };
  if (properties) a.setTableProperties({ story, tableIndex: 0 }, { width: 4000, widthType: 'dxa' });
  else a.insertRow(at, 'below');
  exchange(a, b);
  if (properties) b.setTableProperties({ story, tableIndex: 0 }, { justification: 'center' });
  else b.setColumnWidth(at, 2000);
  expect(a.undo()).toBe(true);
  exchange(a, b);
  expect(b.undo()).toBe(true);
  exchange(a, b);
}

/** The stories that hold the fixture's tables, read from the seeded payloads. */
function locate(target: YrsSession) {
  const cells = (target.tablePayload('body', 0)!.rows as Array<{ cells: Array<{ story: string }> }>)[0]!.cells;
  const controlCell = cells[0]!.story;
  const innerCell = cells[1]!.story;
  const controlStory = target.storySegments(controlCell)
    .find((segment) => segment.kind === 'embed' && segment.embedKind === 'blockSdt');
  const control = controlStory && controlStory.kind === 'embed' ? String(controlStory.payload.story) : '';
  const stories = target.storyIds();
  return {
    control,
    innerCell,
    header: 'hf:top',
    footnote: stories.find((story) => story.startsWith('fn:') && target.tablePayload(story, 0)) ?? '',
    comment: stories.find((story) => story.startsWith('comment:') && target.tablePayload(story, 0)) ?? '',
  };
}

type Where = ReturnType<typeof locate>;

/** Each container kind, damaged the way crossing undos damage it, with the tables both walks must report. */
const CASES: Array<{ name: string; damage: (a: YrsSession, b: YrsSession, where: Where) => void; expected: (where: Where) => YrsTableLoc[] }> = [
  { name: 'no damaged table', damage: () => {}, expected: () => [] },
  { name: 'body', damage: (a, b) => crossing(a, b, 'body'), expected: () => [{ story: 'body', tableIndex: 0 }] },
  { name: 'content control in a cell', damage: (a, b, where) => crossing(a, b, where.control), expected: (where) => [{ story: where.control, tableIndex: 0 }] },
  { name: 'table in a cell', damage: (a, b, where) => crossing(a, b, where.innerCell), expected: (where) => [{ story: where.innerCell, tableIndex: 0 }] },
  { name: 'header with relationship id top', damage: (a, b, where) => crossing(a, b, where.header), expected: (where) => [{ story: where.header, tableIndex: 0 }] },
  { name: 'footnote', damage: (a, b, where) => crossing(a, b, where.footnote), expected: (where) => [{ story: where.footnote, tableIndex: 0 }] },
  { name: 'comment', damage: (a, b, where) => crossing(a, b, where.comment), expected: (where) => [{ story: where.comment, tableIndex: 0 }] },
  { name: 'table that lost only its properties', damage: (a, b) => crossing(a, b, 'body', true), expected: () => [{ story: 'body', tableIndex: 0 }] },
  {
    name: 'inner and outer: only the outer is in reach',
    damage: (a, b, where) => { crossing(a, b, where.innerCell); crossing(a, b, 'body'); },
    expected: () => [{ story: 'body', tableIndex: 0 }],
  },
  {
    name: 'inner under a deleted outer: out of reach',
    damage: (a, b, where) => { crossing(a, b, where.innerCell); crossing(a, b, 'body'); a.deleteTable({ story: 'body', tableIndex: 0 }); exchange(a, b); },
    expected: () => [],
  },
];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))), 60_000);

afterEach(() => {
  for (const created of owned.splice(0)) created.destroy();
});

describe('the engine walk and the projection walk', () => {
  it('the fixture puts a table in every container kind', async () => {
    const probe = await session(60);
    probe.seedFromDocx(containers());
    const where = locate(probe);
    for (const story of [where.control, where.innerCell, where.header, where.footnote, where.comment]) {
      expect(story).not.toBe('');
      expect(probe.tablePayload(story, 0)).not.toBeNull();
    }
  });

  for (const { name, damage, expected } of CASES) {
    it(`report the same tables: ${name}`, async () => {
      const source = containers();
      const a = await session(61);
      const b = await session(62);
      a.seedFromDocx(source);
      b.openDocx(source, false);
      b.applyUpdate(a.encodeState());
      const where = locate(a);
      damage(a, b, where);
      for (const target of [a, b]) {
        const roots = savedRootStories(target, target.materializeDocx()!);
        const engine = sessionUnavailableTablesFrom(target, roots);
        const projection = unavailableTablesByProjection(target, roots);
        expect(engine).toEqual(projection);
        expect(projection).toEqual(expected(where));
      }
    });
  }
});

describe('one scan per save', () => {
  it('a save scans once, an editor save of a fresh projection scans once, and a retained projection is scanned again', async () => {
    const { preloadOpcWasm } = await import('../wasm/opc');
    await preloadOpcWasm();
    const { saveYrsDocx } = await import('./index');
    const { saveEditorDocument } = await import('./editorSave');
    const { yrsToDocument } = await import('./yrsToDocument');
    const { sessionInternals } = await import('./sessionInternals');
    const source = containers();
    const a = await session(61);
    const b = await session(62);
    a.seedFromDocx(source);
    b.openDocx(source, false);
    b.applyUpdate(a.encodeState());
    const internals = sessionInternals(a);
    const real = internals.unavailableTablesFrom.bind(internals);
    let scans = 0;
    internals.unavailableTablesFrom = (roots) => {
      scans += 1;
      return real(roots);
    };

    await saveYrsDocx(a, { updateModifiedDate: false });
    expect(scans).toBe(1);

    scans = 0;
    const projected = yrsToDocument(a, a.materializeDocx()!);
    await saveEditorDocument(a, projected, []);
    expect(scans).toBe(1);

    crossing(a, b, 'body');
    scans = 0;
    let written: unknown = 'not written';
    try {
      written = await saveEditorDocument(a, projected, []);
    } catch (error) {
      expect((error as Error & { code?: string }).code).toBe('table-unavailable');
    }
    expect(written).toBe('not written');
    expect(scans).toBe(1);
  });
});

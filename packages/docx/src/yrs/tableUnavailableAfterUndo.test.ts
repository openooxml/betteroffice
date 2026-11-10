import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createYrsSession, saveYrsDocx, type YrsCellLoc, type YrsSession } from './index';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { saveEditorDocument } from './editorSave';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`;
const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
const CASES = ['S1', 'S2', 'S3'] as const;
type Case = (typeof CASES)[number];

const table = (...cells: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${cells.map(() => '<w:gridCol w:w="4680"/>').join('')}</w:tblGrid><w:tr>${cells.join('')}</w:tr></w:tbl>`;
const TABLE_BODY = `${paragraph('Before')}${table(cell('First'), cell('Second'))}${paragraph('After')}`;
const NESTED_BODY = `${paragraph('Before')}${table(`<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${table(cell('Inner'), cell('Other'))}${paragraph('Outer')}</w:tc>`)}${paragraph('After')}`;

function docx(body = TABLE_BODY): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(RELS));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const at = (row: number, column: number): YrsCellLoc => ({ story: 'body', tableIndex: 0, row, column });
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

/** A and B change the table; A undoes, then B. `synced`: B's change reached A before A's undo. */
async function crossingUndos(kind: Case, synced: boolean) {
  const source = docx();
  const a = await session(61);
  const b = await session(62);
  a.seedFromDocx(source);
  b.openDocx(source, false);
  b.applyUpdate(a.encodeState());
  const original = await savedCopy(source, a.encodeState());
  if (kind === 'S1') a.splitCell(at(0, 1), 1, 2);
  else if (kind === 'S2') a.setCellShading({ anchor: at(0, 0), head: at(0, 0) }, 'FF0000');
  else a.insertRow(at(0, 0), 'below');
  exchange(a, b);
  if (kind === 'S1') b.insertRow(at(0, 0), 'below');
  else if (kind === 'S2') b.setCellShading({ anchor: at(0, 1), head: at(0, 1) }, '00FF00');
  else b.setColumnWidth(at(0, 0), 2000);
  if (synced) exchange(a, b);
  a.undo();
  exchange(a, b);
  const beforeSecondUndo = await savedCopy(source, b.encodeState());
  b.undo();
  exchange(a, b);
  const late = await session(70);
  late.openDocx(source, false);
  late.applyUpdate(a.encodeState());
  return { a, b, late, original, beforeSecondUndo };
}

function payloadKeys(target: YrsSession): string[] {
  const payload = target.tablePayload('body', 0) ?? {};
  return ['tblPr', 'grid', 'rows'].filter((key) => key in payload);
}

async function savedXml(target: YrsSession): Promise<string> {
  const saved = await saveYrsDocx(target, { updateModifiedDate: false });
  return new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
}

/** Saves a throwaway replica holding `update`, so the save leaves the collaborators untouched. */
async function savedCopy(source: Uint8Array, update: Uint8Array): Promise<string> {
  const copy = await createYrsSession({ clientId: 99 });
  try {
    copy.openDocx(source, false);
    copy.applyUpdate(update);
    return await savedXml(copy);
  } finally {
    copy.destroy();
  }
}

const tableXml = (xml: string) => xml.match(/<w:tbl>[\s\S]*<\/w:tbl>/)?.[0] ?? '';
const text = (xml: string) => xml.replace(/<[^>]+>/g, '');
const count = (xml: string, pattern: RegExp) => (xml.match(pattern) ?? []).length;

function caught(call: () => unknown): Error & { code?: string } {
  try {
    call();
  } catch (error) {
    return error as Error & { code?: string };
  }
  throw new Error('expected a refusal');
}

async function caughtAsync(call: () => Promise<unknown>): Promise<Error & { code?: string }> {
  try {
    await call();
  } catch (error) {
    return error as Error & { code?: string };
  }
  throw new Error('expected a refusal');
}

function expectCodedRefusal(error: Error & { code?: string }): void {
  expect(error).toBeInstanceOf(Error);
  expect(error.name).toBe('TableUnavailableError');
  expect(error.code).toBe('table-unavailable');
  expect(error.message).toContain('redo on the collaborator who undid, or delete the table');
}

const TABLE_OPS: Array<[string, (target: YrsSession) => unknown]> = [
  ['insertRow', (target) => target.insertRow(at(0, 0), 'below')],
  ['setCellShading', (target) => target.setCellShading({ anchor: at(0, 0), head: at(0, 0) }, '0000FF')],
  ['setColumnWidth', (target) => target.setColumnWidth(at(0, 0), 3000)],
  ['splitCell', (target) => target.splitCell(at(0, 0), 2, 1)],
];

beforeAll(async () => {
  preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  await preloadOpcWasm();
}, 60_000);

afterEach(() => {
  for (const created of owned.splice(0)) created.destroy();
});

describe('a table that crossing undos emptied', () => {
  for (const kind of CASES) {
    it(`${kind}: every replica and a late joiner project it as an unavailable table the caret cannot enter`, async () => {
      const { a, b, late } = await crossingUndos(kind, false);
      for (const target of [a, b, late]) {
        expect(payloadKeys(target)).toEqual([]);
        const projection = createYrsPositionProjection(target, 'body')!;
        const [table] = projection.allTables();
        expect(table).toMatchObject({ unavailable: true, nodeSize: 2, rowCount: 0, cells: [] });
        expect(projection.tableAtPosition(table!.start + 1)).toBeNull();
        expect(projection.targetAt(table!.start + 1)).toMatchObject({ story: 'body' });
        expect(projection.targetAt(table!.start + 1).cell).toBeUndefined();
        const exported = target.exportStructured({ revisionView: 'accepted' });
        if (!exported.ok) throw new Error('expected an export');
        expect(exported.content.diagnostics).toContainEqual(expect.objectContaining({
          code: 'unsupported-content',
          severity: 'error',
          anchor: { kind: 'table', story: 'body', tableIndex: 0 },
          message: expect.stringContaining('redo on the collaborator who undid, or delete the table'),
        }));
        expect(JSON.stringify(exported.content)).toContain('"element":"w:tbl"');
      }
    });

    it(`${kind}: every save path refuses with a coded error, writes nothing and changes nothing`, async () => {
      const { a, b, late } = await crossingUndos(kind, false);
      for (const target of [a, b, late]) {
        const before = target.encodeState();
        let written: unknown = 'not attempted';
        expectCodedRefusal(await caughtAsync(async () => {
          written = await saveYrsDocx(target, { updateModifiedDate: false });
        }));
        expect(written).toBe('not attempted');
        const base = target.materializeDocx()!;
        expectCodedRefusal(caught(() => yrsToDocument(target, base)));
        expectCodedRefusal(await caughtAsync(async () => {
          written = await saveEditorDocument(target, base, []);
        }));
        expect(written).toBe('not attempted');
        expect(target.encodeState()).toEqual(before);
      }
    });

    it(`${kind}: table operations refuse with a coded error and change nothing`, async () => {
      const { a, b, late } = await crossingUndos(kind, false);
      for (const target of [a, b, late]) {
        const before = target.encodeState();
        for (const [, op] of TABLE_OPS) expectCodedRefusal(caught(() => op(target)));
        expect(target.encodeState()).toEqual(before);
      }
    });

    it(`${kind}: deleting the unavailable table saves the rest of the document unchanged`, async () => {
      const { a, b, original } = await crossingUndos(kind, false);
      a.deleteTable({ story: 'body', tableIndex: 0 });
      exchange(a, b);
      const rest = text(original.replace(tableXml(original), ''));
      expect(rest).toContain('Before');
      expect(rest).toContain('After');
      for (const target of [a, b]) {
        expect(target.tablePayload('body', 0)).toBeNull();
        const xml = await savedXml(target);
        expect(xml).not.toContain('<w:tbl>');
        expect(text(xml)).toBe(rest);
      }
    });

    it(`${kind}: the second undoer's redo restores the table it had and saving works again`, async () => {
      const { a, b, beforeSecondUndo } = await crossingUndos(kind, false);
      const restored = tableXml(beforeSecondUndo);
      expect(count(restored, /<w:tr[ >]/g)).toBeGreaterThan(0);
      expect(text(restored)).toContain('First');
      expect(b.redo()).toBe(true);
      exchange(a, b);
      for (const target of [a, b]) {
        expect(payloadKeys(target)).toEqual(['tblPr', 'grid', 'rows']);
        const table = tableXml(await savedXml(target));
        expect(count(table, /<w:tr[ >]/g)).toBe(count(restored, /<w:tr[ >]/g));
        expect(count(table, /<w:tc>/g)).toBe(count(restored, /<w:tc>/g));
        expect(text(table)).toBe(text(restored));
        expect(createYrsPositionProjection(target, 'body')!.allTables()[0]!.unavailable).toBeUndefined();
      }
    });

    it(`${kind}: undos in synced order keep the table and save it`, async () => {
      const { a, b, late } = await crossingUndos(kind, true);
      for (const target of [a, b, late]) {
        expect(payloadKeys(target)).toEqual(['tblPr', 'grid', 'rows']);
        const table = tableXml(await savedXml(target));
        expect(count(table, /<w:tr[ >]/g)).toBeGreaterThan(0);
        expect(text(table)).toContain('First');
      }
    });
  }
});

async function openPair(source: Uint8Array) {
  const a = await session(61);
  const b = await session(62);
  a.seedFromDocx(source);
  b.openDocx(source, false);
  b.applyUpdate(a.encodeState());
  return { a, b };
}

/** A changes, B changes on top, A undoes before B's change arrives, then B undoes. */
function crossing(a: YrsSession, b: YrsSession, first: (target: YrsSession) => unknown, second: (target: YrsSession) => unknown): void {
  first(a);
  exchange(a, b);
  second(b);
  expect(a.undo()).toBe(true);
  exchange(a, b);
  expect(b.undo()).toBe(true);
  exchange(a, b);
}

/** Display units from the end of the paragraph before the body's first table to the start of the one after it. */
function tableSpan(target: YrsSession): number {
  const [before, after] = target.paragraphs('body');
  const projection = createYrsPositionProjection(target, 'body')!;
  const end = projection.positionForLoc({ story: 'body', paraId: before!.paraId, offset: before!.text.length })!;
  return projection.positionForLoc({ story: 'body', paraId: after!.paraId, offset: 0 })! - end;
}

describe('every save scans the reachable tables', () => {
  it('a session without registered internals still refuses', async () => {
    const { a } = await crossingUndos('S3', false);
    const proxy = new Proxy(a, {}) as YrsSession;
    let written: unknown = 'not attempted';
    expectCodedRefusal(await caughtAsync(async () => {
      written = await saveEditorDocument(proxy, a.materializeDocx()!, []);
    }));
    expectCodedRefusal(await caughtAsync(async () => {
      written = await saveYrsDocx(proxy, { updateModifiedDate: false });
    }));
    expect(written).toBe('not attempted');
  });

  it('undoing the deletion of an unavailable table blocks saving again', async () => {
    const { a } = await crossingUndos('S3', false);
    a.deleteTable({ story: 'body', tableIndex: 0 });
    expect(await savedXml(a)).not.toContain('<w:tbl>');
    expect(a.undo()).toBe(true);
    expect(payloadKeys(a)).toEqual([]);
    expectCodedRefusal(await caughtAsync(() => saveYrsDocx(a, { updateModifiedDate: false })));
  });

  it('a table that lost only its properties takes the same two display units and refuses alike', async () => {
    const { a, b } = await openPair(docx());
    const full = await crossingUndos('S3', false);
    const unavailableSpan = tableSpan(full.a);
    crossing(a, b,
      (target) => target.setTableProperties({ story: 'body', tableIndex: 0 }, { width: 4000, widthType: 'dxa' }),
      (target) => target.setTableProperties({ story: 'body', tableIndex: 0 }, { justification: 'center' }));
    for (const target of [a, b]) {
      expect(payloadKeys(target)).toEqual(['grid', 'rows']);
      const [projected] = createYrsPositionProjection(target, 'body')!.allTables();
      expect(projected).toMatchObject({ unavailable: true, nodeSize: 2, rowCount: 0, cells: [] });
      expect(tableSpan(target)).toBe(unavailableSpan);
      expectCodedRefusal(caught(() => target.insertRow(at(0, 0), 'below')));
      expectCodedRefusal(await caughtAsync(() => saveYrsDocx(target, { updateModifiedDate: false })));
    }
  });

  it('a nested unavailable table blocks saving only while a complete table reaches it', async () => {
    const source = docx(NESTED_BODY);
    const run = async () => {
      const { a, b } = await openPair(source);
      const outerCell = (a.tablePayload('body', 0)!.rows as Array<{ cells: Array<{ story: string }> }>)[0]!.cells[0]!.story;
      const inner = (row: number, column: number): YrsCellLoc => ({ story: outerCell, tableIndex: 0, row, column });
      crossing(a, b, (target) => target.insertRow(inner(0, 0), 'below'), (target) => target.setColumnWidth(inner(0, 0), 2000));
      crossing(a, b, (target) => target.insertRow(at(0, 0), 'below'), (target) => target.setColumnWidth(at(0, 0), 2000));
      expect(payloadKeys(a)).toEqual([]);
      return { a, b, outerCell };
    };

    const deleted = await run();
    expectCodedRefusal(await caughtAsync(() => saveYrsDocx(deleted.a, { updateModifiedDate: false })));
    deleted.a.deleteTable({ story: 'body', tableIndex: 0 });
    exchange(deleted.a, deleted.b);
    for (const target of [deleted.a, deleted.b]) {
      const xml = await savedXml(target);
      expect(xml).not.toContain('<w:tbl>');
      expect(text(xml)).toContain('After');
    }

    const redone = await run();
    expect(redone.b.redo()).toBe(true);
    exchange(redone.a, redone.b);
    for (const target of [redone.a, redone.b]) {
      expect(payloadKeys(target)).toEqual(['tblPr', 'grid', 'rows']);
      const inner = target.tablePayload(redone.outerCell, 0) ?? {};
      expect(['tblPr', 'grid', 'rows'].filter((key) => key in inner)).toEqual([]);
      expectCodedRefusal(await caughtAsync(() => saveYrsDocx(target, { updateModifiedDate: false })));
    }
    redone.a.deleteTable({ story: redone.outerCell, tableIndex: 0 });
    exchange(redone.a, redone.b);
    for (const target of [redone.a, redone.b]) {
      const xml = await savedXml(target);
      expect(count(xml, /<w:tbl>/g)).toBe(1);
      expect(text(xml)).toContain('Outer');
    }
  });
});

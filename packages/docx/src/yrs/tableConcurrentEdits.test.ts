import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsCellLoc, type YrsSession } from './index';
import { createYrsPositionProjection } from './yrsPositionProjection';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const sessions: YrsSession[] = [];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))), 30_000);
afterEach(() => { for (const session of sessions.splice(0)) session.destroy(); });

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
const body = paragraph('Before') +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid><w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>` +
  paragraph('After');

function docx(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(clientId: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(docx(), true);
  session.beginUndoCapture();
  return session;
}

async function peerOf(source: YrsSession, clientId: number): Promise<YrsSession> {
  const peer = await createYrsSession({ clientId });
  sessions.push(peer);
  peer.openDocx(docx(), false);
  peer.loadState(source.encodeState());
  peer.beginUndoCapture();
  return peer;
}

function sync(a: YrsSession, b: YrsSession): void {
  a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
  b.applyUpdate(a.encodeStateAsUpdate(b.encodeStateVector()));
}

const at = (row: number, column: number): YrsCellLoc => ({ story: 'body', tableIndex: 0, row, column });

/** Cell stories that a table of a live story names but that do not exist. */
function danglingCells(session: YrsSession): string[] {
  const stories = session.storyIds();
  const missing: string[] = [];
  for (const story of stories) {
    let segments;
    try { segments = session.storySegments(story); } catch { continue; }
    for (const segment of segments) {
      if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
      for (const row of (segment.payload.rows as { cells?: { story?: string }[] }[] | undefined) ?? []) {
        for (const named of row.cells ?? []) {
          if (named.story && !stories.includes(named.story)) missing.push(named.story);
        }
      }
    }
  }
  return missing;
}

const message = (error: unknown) => String((error as Error)?.message ?? error).slice(0, 160);

async function effects(session: YrsSession, clientId: number) {
  const report: Record<string, unknown> = { dangling: danglingCells(session) };
  try { createYrsPositionProjection(session, 'body'); report.cellLookup = 'ok'; } catch (error) { report.cellLookup = `throws: ${message(error)}`; }
  try { session.insertRow(at(0, 0), 'below'); report.tableCommand = 'ok'; } catch (error) { report.tableCommand = `throws: ${message(error)}`; }
  try {
    const saved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    report.save = `ok ${saved.length} bytes`;
    const reopened = await createYrsSession({ clientId });
    sessions.push(reopened);
    reopened.openDocx(saved, true);
    report.reopenDangling = danglingCells(reopened);
    try { createYrsPositionProjection(reopened, 'body'); report.reopenCellLookup = 'ok'; } catch (error) { report.reopenCellLookup = `throws: ${message(error)}`; }
  } catch (error) {
    report.save = `throws: ${message(error)}`;
  }
  return report;
}

/** Every live story's segments, sorted by story: equal across converged replicas. */
function content(session: YrsSession): string {
  return JSON.stringify(session.storyIds().sort().map((story) => [story, session.storySegments(story)]));
}

function syncAll(peers: YrsSession[]): void {
  for (let round = 0; round < 2; round++) for (const a of peers) for (const b of peers) if (a !== b) a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
}

function expectConverged(peers: YrsSession[]): void {
  const [first, ...rest] = peers;
  for (const peer of rest) {
    expect(content(peer)).toBe(content(first!));
    expect(Array.from(peer.encodeStateVector())).toEqual(Array.from(first!.encodeStateVector()));
  }
}

function rows(session: YrsSession): number {
  return ((session.tablePayload('body', 0)?.rows as unknown[] | undefined) ?? []).length;
}

async function expectSound(session: YrsSession, clientId: number) {
  const report = await effects(session, clientId);
  console.log(`EFFECTS ${JSON.stringify(report)}`);
  expect(report).toMatchObject({ dangling: [], cellLookup: 'ok', reopenDangling: [], reopenCellLookup: 'ok' });
}

test('two peers: undoing a split after the other peer rewrote the table keeps the cell that table names', async () => {
  const a = await open(9701);
  const b = await peerOf(a, 9702);
  a.splitCell(at(0, 1), 1, 2);
  sync(a, b);
  const created = a.storyIds().filter((story) => story.includes('~9701'));
  expect(created.length).toBeGreaterThan(0);
  b.insertRow(at(0, 0), 'below');
  sync(a, b);
  expect(a.undo()).toBe(true);
  sync(a, b);
  expect({ a: danglingCells(a), b: danglingCells(b) }).toEqual({ a: [], b: [] });
  expect(rows(a)).toBe(2);
  for (const story of created) expect(a.storyIds()).toContain(story);
  expectConverged([a, b]);
  expect(a.redo()).toBe(true);
  sync(a, b);
  expect({ a: danglingCells(a), b: danglingCells(b) }).toEqual({ a: [], b: [] });
  expectConverged([a, b]);
  await expectSound(a, 9703);
});

test('three peers: the undo after a remote table rewrite converges on every peer', async () => {
  const a = await open(9731);
  const b = await peerOf(a, 9732);
  const c = await peerOf(a, 9733);
  a.splitCell(at(0, 1), 1, 2);
  syncAll([a, b, c]);
  b.insertRow(at(0, 0), 'below');
  syncAll([a, b, c]);
  expect(a.undo()).toBe(true);
  syncAll([a, b, c]);
  for (const peer of [a, b, c]) expect(danglingCells(peer)).toEqual([]);
  expectConverged([a, b, c]);
  await expectSound(c, 9734);
});

test('two peers: re-splitting after the other peer removed the first split cell leaves no dangling cell', async () => {
  const a = await open(9711);
  const b = await peerOf(a, 9712);
  a.splitCell(at(0, 1), 1, 2);
  sync(a, b);
  b.deleteColumn({ anchor: at(0, 2), head: at(0, 2) });
  sync(a, b);
  a.undo();
  sync(a, b);
  a.splitCell(at(0, 1), 1, 2);
  sync(a, b);
  expect({ a: danglingCells(a), b: danglingCells(b) }).toEqual({ a: [], b: [] });
  expectConverged([a, b]);
  await expectSound(a, 9713);
});

for (const [name, steps] of [
  ['split, insert row, undo twice, redo twice', (s: YrsSession) => { s.splitCell(at(0, 1), 1, 2); s.insertRow(at(0, 0), 'below'); s.undo(); s.undo(); s.redo(); s.redo(); }],
  ['split, delete the new column, undo, undo, split again', (s: YrsSession) => { s.splitCell(at(0, 1), 1, 2); s.deleteColumn({ anchor: at(0, 2), head: at(0, 2) }); s.undo(); s.undo(); s.splitCell(at(0, 1), 1, 2); }],
  ['split, merge back, undo, split again', (s: YrsSession) => { s.splitCell(at(0, 1), 1, 2); s.mergeCells({ anchor: at(0, 1), head: at(0, 2) }); s.undo(); s.splitCell(at(0, 0), 1, 2); }],
] as const) {
  test(`one user: ${name} leaves no dangling cell`, async () => {
    const a = await open(9721);
    const errors: string[] = [];
    try { steps(a); } catch (error) { errors.push(message(error)); }
    const report = { dangling: danglingCells(a), errors };
    console.log(`SCENARIO one-user ${name}: ${JSON.stringify(report)}`);
    expect(report.dangling).toEqual([]);
  });
}

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

function docx(content = body): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${content}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function open(clientId: number, content = body): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(docx(content), true);
  session.beginUndoCapture();
  return session;
}

async function peerOf(source: YrsSession, clientId: number, content = body): Promise<YrsSession> {
  const peer = await createYrsSession({ clientId });
  sessions.push(peer);
  peer.openDocx(docx(content), false);
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

function expectConverged(peers: YrsSession[]): void {
  const [first, ...rest] = peers;
  for (const peer of rest) {
    expect(content(peer)).toBe(content(first!));
    expect(Array.from(peer.encodeStateVector())).toEqual(Array.from(first!.encodeStateVector()));
  }
}

async function expectSound(session: YrsSession, clientId: number) {
  const report = await effects(session, clientId);
  console.log(`EFFECTS ${JSON.stringify(report)}`);
  expect(report).toMatchObject({ dangling: [], cellLookup: 'ok', reopenDangling: [], reopenCellLookup: 'ok' });
}

function syncAll(peers: YrsSession[]): void {
  for (let round = 0; round < 2; round++) for (const a of peers) for (const b of peers) if (a !== b) a.applyUpdate(b.encodeStateAsUpdate(a.encodeStateVector()));
}

function rows(session: YrsSession): number {
  return ((session.tablePayload('body', 0)?.rows as unknown[] | undefined) ?? []).length;
}

function cellText(session: YrsSession, story: string): string | null {
  try { return session.paragraphs(story).map((p) => p.text).join('\n'); } catch { return null; }
}

const FONT = resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const REGION_REQUEST = JSON.stringify({
  bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } }, renderEnv: {},
});

function render(session: YrsSession) {
  const row: Record<string, unknown> = {};
  try {
    const layout = JSON.parse(session.layoutDocumentWithRegionsRetainedJson(REGION_REQUEST)) as { pages?: unknown[] };
    row.layout = `ok pages=${Array.isArray(layout.pages) ? layout.pages.length : '?'}`;
  } catch (error) { row.layout = `throws: ${message(error)}`; }
  try { row.frame = `ok ${session.buildDisplayListFrame('{}', 0).length} bytes`; } catch (error) { row.frame = `throws: ${message(error)}`; }
  try { createYrsPositionProjection(session, 'body'); row.cellLookup = 'ok'; } catch (error) { row.cellLookup = `throws: ${message(error)}`; }
  return row;
}

test('severity probe: layout and display in the dangling state on every peer', async () => {
  const a = await open(9761);
  const b = await peerOf(a, 9762);
  const c = await peerOf(a, 9763);
  for (const peer of [a, b, c]) peer.registerFont(new Uint8Array(readFileSync(FONT)));
  const before = { a: render(a), b: render(b), c: render(c) };
  a.splitCell(at(0, 1), 1, 2);
  syncAll([a, b, c]);
  b.insertRow(at(0, 0), 'below');
  syncAll([a, b, c]);
  a.undo();
  syncAll([a, b, c]);
  const after = { a: render(a), b: render(b), c: render(c), dangling: danglingCells(a) };
  console.log(`SEVERITY-PROBE ${JSON.stringify({ before, after })}`);
  expect(after.dangling.length).toBeGreaterThan(0);
});

test('two peers: undoing a split deletes the text the other collaborator typed into the new cell (fact probe)', async () => {
  const a = await open(9751);
  const b = await peerOf(a, 9752);
  a.splitCell(at(0, 1), 1, 2);
  sync(a, b);
  const created = a.storyIds().filter((story) => story.includes('~9751'));
  expect(created.length).toBeGreaterThan(0);
  const target = created[0]!;
  const [para] = b.paragraphs(target);
  b.insertText({ story: target, paraId: para!.paraId, offset: 0 }, 'typed by B');
  sync(a, b);
  const before = { a: cellText(a, target), b: cellText(b, target) };
  expect(a.undo()).toBe(true);
  sync(a, b);
  const after = { a: cellText(a, target), b: cellText(b, target), aStories: a.storyIds().includes(target), bStories: b.storyIds().includes(target),
    aAnywhere: a.storyIds().some((story) => (cellText(a, story) ?? '').includes('typed by B')), bAnywhere: b.storyIds().some((story) => (cellText(b, story) ?? '').includes('typed by B')),
    dangling: { a: danglingCells(a), b: danglingCells(b) } };
  console.log(`CONTENT-LOSS-PROBE ${JSON.stringify({ target, before, after })}`);
  expect(before).toEqual({ a: 'typed by B', b: 'typed by B' });
});

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


async function frozenRoom(base: number) {
  const a = await open(base + 1);
  const b = await peerOf(a, base + 2);
  const c = await peerOf(a, base + 3);
  for (const peer of [a, b, c]) peer.registerFont(new Uint8Array(readFileSync(FONT)));
  a.splitCell(at(0, 1), 1, 2);
  syncAll([a, b, c]);
  b.insertRow(at(0, 0), 'below');
  syncAll([a, b, c]);
  a.undo();
  syncAll([a, b, c]);
  return { a, b, c };
}

function attempt(label: string, run: () => unknown): string {
  try { const value = run(); return `${label} ok${value === undefined ? '' : ` ${JSON.stringify(value).slice(0, 80)}`}`; } catch (error) { return `${label} throws: ${message(error)}`; }
}

test('permanence probe: a late joiner and un-freeze attempts in the dangling state', async () => {
  const rows: Record<string, unknown> = {};
  {
    const { a } = await frozenRoom(9770);
    const d = await createYrsSession({ clientId: 9779 });
    sessions.push(d);
    const steps: string[] = [];
    steps.push(attempt('open', () => { d.openDocx(docx(), false); d.loadState(a.encodeState()); d.beginUndoCapture(); }));
    d.registerFont(new Uint8Array(readFileSync(FONT)));
    steps.push(`dangling ${JSON.stringify(danglingCells(d))}`);
    const first = render(d);
    steps.push(`first ${JSON.stringify(first)}`);
    steps.push(attempt('editBody', () => { const [p] = d.paragraphs('body'); d.insertText({ story: 'body', paraId: p!.paraId, offset: 0 }, 'D'); }));
    steps.push(`afterEdit ${JSON.stringify(render(d))}`);
    try {
      const saved = (await saveYrsDocx(d, { updateModifiedDate: false })).bytes;
      const reopened = await createYrsSession({ clientId: 9778 });
      sessions.push(reopened);
      reopened.openDocx(saved, true);
      reopened.registerFont(new Uint8Array(readFileSync(FONT)));
      steps.push(`save ok ${saved.length} bytes; reopened dangling ${JSON.stringify(danglingCells(reopened))} ${JSON.stringify(render(reopened))}`);
    } catch (error) { steps.push(`save throws: ${message(error)}`); }
    const fresh = await createYrsSession({ clientId: 9777 });
    sessions.push(fresh);
    steps.push(attempt('openFromStateOnly', () => { fresh.loadState(a.encodeState()); }));
    fresh.registerFont(new Uint8Array(readFileSync(FONT)));
    steps.push(`stateOnly ${JSON.stringify(render(fresh))}`);
    rows.lateJoiner = steps;
  }
  const unfreeze: Array<[string, (room: { a: YrsSession; b: YrsSession; c: YrsSession }) => unknown]> = [
    ['A undo again', ({ a }) => a.undo()],
    ['A redo', ({ a }) => a.redo()],
    ['B undo', ({ b }) => b.undo()],
    ['B deleteTable', ({ b }) => b.deleteTable({ story: 'body', tableIndex: 0 })],
    ['B insertRow again', ({ b }) => b.insertRow(at(0, 0), 'below')],
    ['C insertRow', ({ c }) => c.insertRow(at(1, 0), 'below')],
  ];
  let base = 9780;
  for (const [label, run] of unfreeze) {
    const room = await frozenRoom(base);
    base += 10;
    const step = attempt(label, () => run(room));
    syncAll([room.a, room.b, room.c]);
    rows[label] = { step, a: render(room.a), b: render(room.b), c: render(room.c), dangling: danglingCells(room.a) };
  }
  console.log(`PERMANENCE-PROBE ${JSON.stringify(rows)}`);
  expect(Object.keys(rows).length).toBe(unfreeze.length + 1);
}, 120_000);

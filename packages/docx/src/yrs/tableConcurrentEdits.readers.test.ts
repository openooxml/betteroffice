import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { createYrsSidebarProjection } from '../layout/render/yrsSidebarProjection';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsCellLoc, type YrsLoc, type YrsSession } from './index';
import { readTableFragment } from './clipboardFragment';
import { storyPlainText } from './storyPlainText';
import { tableBoundaryDeleteTarget } from './tableBoundaryDelete';
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

type Order = 'synced' | 'concurrent';

async function danglingRoom(base: number, count: 2 | 3, order: Order) {
  const a = await open(base + 1);
  const peers = [a];
  for (let index = 2; index <= count; index++) peers.push(await peerOf(a, base + index));
  for (const peer of peers) peer.registerFont(new Uint8Array(readFileSync(FONT)));
  const b = peers[1]!;
  a.splitCell(at(0, 1), 1, 2);
  syncAll(peers);
  b.insertRow(at(0, 0), 'below');
  if (order === 'synced') syncAll(peers);
  expect(a.undo()).toBe(true);
  syncAll(peers);
  return { a, b, peers };
}

function firstParagraph(session: YrsSession, story: string): YrsLoc {
  return { story, paraId: session.paragraphSpans(story)[0]!.paraId, offset: 0 };
}

/** Every reader of a table cell, run against the session; `ok` or the error each throws. */
function readerRows(session: YrsSession): Record<string, string> {
  const rows: Record<string, string> = {};
  const check = (name: string, run: () => unknown) => {
    try { run(); rows[name] = 'ok'; } catch (error) { rows[name] = `throws: ${message(error)}`; }
  };
  check('projection', () => createYrsPositionProjection(session, 'body'));
  check('layout', () => session.layoutDocumentWithRegionsRetainedJson(REGION_REQUEST));
  check('frame', () => session.buildDisplayListFrame('{}', 0));
  check('plainText', () => storyPlainText(session, 'body'));
  check('sidebar', () => createYrsSidebarProjection(session));
  if (session.tablePayload('body', 0)) check('copyTable', () => readTableFragment(session, { story: 'body', tableIndex: 0 }));
  for (const story of danglingCells(session)) {
    check(`${story} paragraphs`, () => {
      const paragraphs = session.paragraphs(story);
      if (paragraphs.length !== 1 || paragraphs[0]!.text !== '') throw new Error(JSON.stringify(paragraphs));
    });
    check(`${story} hasStory`, () => { if (!session.hasStory(story)) throw new Error('hasStory is false'); });
    check(`${story} boundaryDelete`, () => tableBoundaryDeleteTarget({
      paragraphSpans: (id) => session.paragraphSpans(id),
      paragraphSegments: (id, indices) => session.storySegmentUnits(id, indices),
    }, firstParagraph(session, story), 'backward'));
    check(`${story} sticky`, () => {
      const loc = session.resolveStickyPosition(session.encodeStickyPosition(firstParagraph(session, story)));
      if (loc?.story !== story || loc.offset !== 0) throw new Error(JSON.stringify(loc));
    });
  }
  return rows;
}

function expectReadable(session: YrsSession): void {
  for (const [reader, outcome] of Object.entries(readerRows(session))) expect(`${reader}: ${outcome}`).toBe(`${reader}: ok`);
}

function typeInto(session: YrsSession, story: string, text: string): void {
  session.insertText(firstParagraph(session, story), text);
}

async function expectSavedSound(session: YrsSession, clientId: number): Promise<void> {
  const saved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
  const reopened = await createYrsSession({ clientId });
  sessions.push(reopened);
  reopened.openDocx(saved, true);
  expect(danglingCells(reopened)).toEqual([]);
  createYrsPositionProjection(reopened, 'body');
}

for (const order of ['synced', 'concurrent'] as const) {
  for (const count of [2, 3] as const) {
    test(`${count} peers, ${order}: a named-but-missing cell reads as one empty paragraph and typing into it converges`, async () => {
      const { a, b, peers } = await danglingRoom(9800 + (order === 'synced' ? 0 : 50) + count * 10, count, order);
      const missing = danglingCells(a);
      console.log(`READERS ${order} ${count} ${JSON.stringify({ missing })}`);
      if (order === 'synced') expect(missing.length).toBeGreaterThan(0);
      for (const peer of peers) expectReadable(peer);
      const before = Array.from(a.encodeStateVector());
      for (const peer of peers) readerRows(peer);
      expect(Array.from(a.encodeStateVector())).toEqual(before);
      b.insertRow(at(0, 0), 'below');
      syncAll(peers);
      for (const peer of peers) expectReadable(peer);
      for (const story of missing) {
        typeInto(peers[peers.length - 1]!, story, `typed ${story}`);
        syncAll(peers);
        for (const peer of peers) expect(cellText(peer, story)).toBe(`typed ${story}`);
      }
      for (const peer of peers) expect(danglingCells(peer)).toEqual([]);
      expectConverged(peers);
      await expectSavedSound(a, 9899);
    });
  }
}

test('typing into a missing cell is one undo step: undo leaves an empty real cell and redo restores the text', async () => {
  const { a, b } = await danglingRoom(9900, 2, 'synced');
  const [story] = danglingCells(b);
  expect(story).toBeDefined();
  expect(rows(b)).toBe(2);
  typeInto(b, story!, 'x');
  sync(a, b);
  expect(b.undo()).toBe(true);
  sync(a, b);
  for (const peer of [a, b]) {
    expect(cellText(peer, story!)).toBe('');
    expect(danglingCells(peer)).toEqual([]);
    expectReadable(peer);
  }
  expect(rows(b)).toBe(2);
  expect(b.redo()).toBe(true);
  sync(a, b);
  expect(cellText(a, story!)).toBe('x');
  expect(b.undo()).toBe(true);
  expect(b.undo()).toBe(true);
  sync(a, b);
  expect(rows(b)).toBe(1);
  for (const peer of [a, b]) expectReadable(peer);
  expectConverged([a, b]);
});

test('two collaborators type into the same missing cell before syncing: one creation is kept (named limit)', async () => {
  const { a, b } = await danglingRoom(9920, 2, 'synced');
  const [story] = danglingCells(a);
  typeInto(a, story!, 'from A');
  typeInto(b, story!, 'from B');
  sync(a, b);
  sync(a, b);
  const text = cellText(a, story!);
  console.log(`CONCURRENT-FIRST-WRITES ${JSON.stringify({ text, entries: a.storyIds().filter((id) => id === story).length })}`);
  expect(['from A', 'from B']).toContain(text);
  for (const peer of [a, b]) {
    expect(danglingCells(peer)).toEqual([]);
    expectReadable(peer);
  }
  expectConverged([a, b]);
  typeInto(a, story!, 'next ');
  sync(a, b);
  expect(cellText(b, story!)).toBe(`next ${text}`);
});

test('placing a caret in a missing cell and reading it writes nothing; a restored caret resolves without writing', async () => {
  const { a, b } = await danglingRoom(9940, 2, 'synced');
  const [story] = danglingCells(b);
  const before = Array.from(b.encodeStateVector());
  const caret = firstParagraph(b, story!);
  b.setSelection(caret, caret);
  expect(b.selection()).toMatchObject({ anchor: { story, offset: 0 }, head: { story, offset: 0 } });
  const sticky = b.encodeStickyPosition(caret);
  expect(b.resolveStickyPosition(sticky)).toMatchObject({ story, offset: 0 });
  expectReadable(b);
  expect(Array.from(b.encodeStateVector())).toEqual(before);

  const restored = await peerOf(b, 9949);
  restored.setSelection(caret, caret);
  expect(restored.selection()).toMatchObject({ anchor: { story, offset: 0 } });
  expect(restored.resolveStickyPosition(sticky)).toMatchObject({ story, offset: 0 });
  expect(Array.from(restored.encodeStateVector())).toEqual(before);

  typeInto(a, story!, 'later');
  sync(a, b);
  expect(b.selection()).toMatchObject({ anchor: { story, offset: 0 } });
  expect(cellText(b, story!)).toBe('later');
});

function cellOf(session: YrsSession, story: string): YrsCellLoc {
  const payload = session.tablePayload('body', 0) as { rows?: { cells?: { story?: string }[] }[] } | null;
  for (const [row, value] of (payload?.rows ?? []).entries()) {
    const column = (value.cells ?? []).findIndex((cell) => cell.story === story);
    if (column >= 0) return at(row, column);
  }
  throw new Error(`no cell names ${story}`);
}

const collapsed = (loc: YrsLoc) => ({ story: loc.story, start: { paraId: loc.paraId, offset: 0 }, end: { paraId: loc.paraId, offset: 0 } });

/** Every write entry point a user or host can aim at one cell. */
const WRITES: Array<[string, (session: YrsSession, story: string) => unknown]> = [
  ['insertText', (s, story) => s.insertText(firstParagraph(s, story), 'typed')],
  ['replaceRange', (s, story) => s.replaceRange(collapsed(firstParagraph(s, story)), 'pasted')],
  ['deleteRange at start', (s, story) => s.deleteRange(collapsed(firstParagraph(s, story)))],
  ['mergeParagraphs (backspace at start)', (s, story) => s.mergeParagraphs(story, firstParagraph(s, story).paraId)],
  ['insertLineBreak', (s, story) => s.insertLineBreak(firstParagraph(s, story))],
  ['splitParagraph', (s, story) => s.splitParagraph(firstParagraph(s, story))],
  ['toggleMark', (s, story) => s.toggleMark(collapsed(firstParagraph(s, story)), { type: 'bold' })],
  ['applyParagraphStyle', (s, story) => s.applyParagraphStyle(collapsed(firstParagraph(s, story)), 'Heading1')],
  ['setHyperlink', (s, story) => s.setHyperlink(collapsed(firstParagraph(s, story)), { href: 'https://example.com' })],
  ['insertTable', (s, story) => s.insertTable(firstParagraph(s, story), 1, 1)],
  ['insertPageBreak', (s, story) => s.insertPageBreak(firstParagraph(s, story))],
  ['addComment', (s, story) => s.addComment([collapsed(firstParagraph(s, story))], 'Reviewer', '2026-01-01T00:00:00Z', 'note')],
  ['mergeCells with a neighbour', (s, story) => { const cell = cellOf(s, story); return s.mergeCells({ anchor: { ...cell, column: Math.max(0, cell.column - 1) }, head: cell }); }],
  ['splitCell', (s, story) => s.splitCell(cellOf(s, story), 1, 2)],
  ['deleteRow', (s, story) => { const cell = cellOf(s, story); return s.deleteRow({ anchor: cell, head: cell }); }],
  ['deleteColumn', (s, story) => { const cell = cellOf(s, story); return s.deleteColumn({ anchor: cell, head: cell }); }],
  ['deleteTable', (s) => s.deleteTable({ story: 'body', tableIndex: 0 })],
  ['host batch insertParagraphs', (s, story) => s.applyEdits({
    expectVersion: s.version(),
    steps: [{ op: 'insertParagraphs', target: { story, paraId: firstParagraph(s, story).paraId }, at: 'end', paragraphs: [{ text: 'from host' }] }],
  })],
];

test('every write entry point aimed at a missing cell works or refuses cleanly, and the room stays readable and converges', async () => {
  const outcomes: Record<string, string> = {};
  const problems: string[] = [];
  let base = 10_000;
  for (const [name, write] of WRITES) {
    const { a, b, peers } = await danglingRoom(base, 2, 'synced');
    base += 10;
    const [story] = danglingCells(b);
    expect(story).toBeDefined();
    const before = Array.from(b.encodeStateVector());
    const outcome = attempt(name, () => write(b, story!));
    const wrote = JSON.stringify(Array.from(b.encodeStateVector())) !== JSON.stringify(before);
    outcomes[name] = `${outcome}${wrote ? ' (wrote)' : ''}`;
    if (/unreachable|panicked|RuntimeError/.test(outcome)) problems.push(`${name}: engine trap ${outcome}`);
    if (outcome.includes(' throws: ') && wrote && JSON.stringify(b.paragraphs(story!).map((paragraph) => paragraph.text)) !== '[""]') problems.push(`${name}: refused after writing content`);
    syncAll(peers);
    for (const peer of [a, b]) {
      for (const [reader, result] of Object.entries(readerRows(peer))) if (result !== 'ok') problems.push(`${name}: ${reader} ${result}`);
    }
    if (content(a) !== content(b)) problems.push(`${name}: peers diverge`);
  }
  console.log(`WRITE-SURFACE ${JSON.stringify(outcomes)}`);
  console.log(`WRITE-SURFACE-PROBLEMS ${JSON.stringify(problems)}`);
  expect(problems).toEqual([]);
}, 120_000);

test('the first write into a removed cell that a live table still names holds only the new text, on both peers and after save and reopen', async () => {
  const a = await open(10_300);
  const b = await peerOf(a, 10_301);
  const story = (a.tablePayload('body', 0) as { rows: { cells: { story: string }[] }[] }).rows[0]!.cells[1]!.story;
  expect(cellText(a, story)).toBe('B');
  a.deleteStory(story);
  sync(a, b);
  a.insertRow(at(0, 0), 'below');
  expect(a.undo()).toBe(true);
  sync(a, b);
  for (const peer of [a, b]) {
    expect(peer.storyIds()).not.toContain(story);
    expect(peer.paragraphs(story).map((paragraph) => paragraph.text)).toEqual(['']);
    expectReadable(peer);
  }
  typeInto(b, story, 'new');
  sync(a, b);
  for (const peer of [a, b]) {
    expect(peer.paragraphs(story).map((paragraph) => paragraph.text)).toEqual(['new']);
    expect(peer.storyIds().filter((id) => id === story)).toEqual([story]);
  }
  expectConverged([a, b]);
  const saved = (await saveYrsDocx(a, { updateModifiedDate: false })).bytes;
  const reopened = await createYrsSession({ clientId: 10_302 });
  sessions.push(reopened);
  reopened.openDocx(saved, true);
  const texts = reopened.storyIds().filter((id) => id.includes(':t')).map((id) => cellText(reopened, id) ?? '');
  expect(texts.some((text) => text === 'new')).toBe(true);
  expect(texts.some((text) => text.includes('B'))).toBe(false);
});

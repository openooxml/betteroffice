import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { createYrsSidebarProjection } from '../layout/render/yrsSidebarProjection';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { createEditSession, preloadEditWasm, type EditSession } from '../wasm/edit';
import { createYrsSession, saveYrsDocx, type YrsAuthor, type YrsCellLoc, type YrsLoc, type YrsSession, type YrsStoryRange } from './index';
import { readTableFragment } from './clipboardFragment';
import type { DocxWorkerCommandEngine } from './residentCommandProtocol';
import { resolveLiveCell } from './residentCommandTableSelection';
import { storyPlainText } from './storyPlainText';
import { tableBoundaryDeleteTarget } from './tableBoundaryDelete';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { wrapSession } from './yrsSessionFacade';

const path = (relative: string) => fileURLToPath(new URL(relative, import.meta.url));
const WASM = path('../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = path('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const sessions: YrsSession[] = [];
const natives = new WeakMap<YrsSession, EditSession>();

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))), 30_000);
afterEach(() => { for (const session of sessions.splice(0)) session.destroy(); });

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
const body = paragraph('Before') +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid><w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>` +
  paragraph('After');
const AUTHOR: YrsAuthor = { name: 'Reviewer', date: '2026-01-01T00:00:00Z' };

function docx(content = body): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${content}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const font = () => new Uint8Array(readFileSync(FONT));

/** A session whose raw engine stays reachable, for entry points the facade does not expose. */
function createSession(clientId: number): YrsSession {
  const native = createEditSession(clientId);
  const session = wrapSession(native, clientId);
  natives.set(session, native);
  sessions.push(session);
  return session;
}

async function open(clientId: number, content = body): Promise<YrsSession> {
  const session = createSession(clientId);
  session.openDocx(docx(content), true);
  session.beginUndoCapture();
  return session;
}

async function peerOf(source: YrsSession, clientId: number, content = body): Promise<YrsSession> {
  const peer = createSession(clientId);
  peer.openDocx(docx(content), false);
  peer.loadState(source.encodeState());
  peer.beginUndoCapture();
  return peer;
}

/** A replica that joins late from the full update of `source`. */
function lateJoiner(source: YrsSession, clientId: number): YrsSession {
  const joiner = createSession(clientId);
  joiner.openDocx(docx(), false);
  joiner.applyUpdate(source.encodeStateAsUpdate());
  joiner.registerFont(font());
  return joiner;
}

/** A session opened from encoded state alone, with no package. */
function stateOnly(source: YrsSession, clientId: number): YrsSession {
  const session = createSession(clientId);
  session.loadState(source.encodeState());
  session.registerFont(font());
  return session;
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

const fullMessage = (error: unknown) => String((error as Error)?.message ?? error);
const message = (error: unknown) => fullMessage(error).slice(0, 160);

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

/** The document state of one replica as bytes: its state vector and its full update. */
function stateBytes(session: YrsSession): string {
  return `${Array.from(session.encodeStateVector()).join(',')}|${Array.from(session.encodeStateAsUpdate()).join(',')}`;
}

function cellText(session: YrsSession, story: string): string | null {
  try { return session.paragraphs(story).map((p) => p.text).join('\n'); } catch { return null; }
}

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
  for (const peer of [a, b, c]) peer.registerFont(font());
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
  for (const peer of [a, b, c]) peer.registerFont(font());
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
    d.registerFont(font());
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
      reopened.registerFont(font());
      steps.push(`save ok ${saved.length} bytes; reopened dangling ${JSON.stringify(danglingCells(reopened))} ${JSON.stringify(render(reopened))}`);
    } catch (error) { steps.push(`save throws: ${message(error)}`); }
    const fresh = await createYrsSession({ clientId: 9777 });
    sessions.push(fresh);
    steps.push(attempt('openFromStateOnly', () => { fresh.loadState(a.encodeState()); }));
    fresh.registerFont(font());
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

/**
 * A splits a cell, B inserts a row, A undoes the split. `synced` exchanges B's row before the
 * undo; `concurrent` lets A undo while B inserts the row, with no sync between.
 */
async function danglingRoom(base: number, count: 2 | 3, order: Order) {
  const a = await open(base + 1);
  const peers = [a];
  for (let index = 2; index <= count; index++) peers.push(await peerOf(a, base + index));
  for (const peer of peers) peer.registerFont(font());
  const b = peers[1]!;
  a.splitCell(at(0, 1), 1, 2);
  syncAll(peers);
  b.insertRow(at(0, 0), 'below');
  if (order === 'synced') syncAll(peers);
  expect(a.undo()).toBe(true);
  syncAll(peers);
  const missing = danglingCells(a);
  expect(missing.length).toBeGreaterThan(0);
  return { a, b, peers, story: missing[0]! };
}

function firstParagraph(session: YrsSession, story: string): YrsLoc {
  return { story, paraId: session.paragraphSpans(story)[0]!.paraId, offset: 0 };
}

const collapsed = (loc: YrsLoc): YrsStoryRange => ({ story: loc.story, start: { paraId: loc.paraId, offset: loc.offset }, end: { paraId: loc.paraId, offset: loc.offset } });

function cellOf(session: YrsSession, story: string): YrsCellLoc {
  const payload = session.tablePayload('body', 0) as { rows?: { cells?: { story?: string }[] }[] } | null;
  for (const [row, value] of (payload?.rows ?? []).entries()) {
    const column = (value.cells ?? []).findIndex((entry) => entry.story === story);
    if (column >= 0) return at(row, column);
  }
  throw new Error(`no cell names ${story}`);
}

const commandReader = (session: YrsSession) => ({ geometryReader: session }) as unknown as DocxWorkerCommandEngine;

function expectOneEmptyParagraph(story: string, paragraphs: ReadonlyArray<{ paraId: string; text: string }>): void {
  if (paragraphs.length !== 1 || paragraphs[0]!.text !== '' || paragraphs[0]!.paraId !== `${story}#p0`) {
    throw new Error(JSON.stringify(paragraphs));
  }
}

/** Every reader of the document and of each missing cell, run against the session; `ok` or the error each throws. */
function readerRows(session: YrsSession): Record<string, string> {
  const rows: Record<string, string> = {};
  const check = (name: string, run: () => unknown) => {
    try { run(); rows[name] = 'ok'; } catch (error) { rows[name] = `throws: ${message(error)}`; }
  };
  check('projection', () => createYrsPositionProjection(session, 'body')?.allTables());
  check('layout', () => session.layoutDocumentWithRegionsRetainedJson(REGION_REQUEST));
  check('frame', () => session.buildDisplayListFrame('{}', 0));
  check('plainText', () => storyPlainText(session, 'body'));
  check('sidebar', () => createYrsSidebarProjection(session));
  if (session.tablePayload('body', 0)) check('copyTable', () => readTableFragment(session, { story: 'body', tableIndex: 0 }));
  check('searchText', () => session.searchText('A'));
  check('findText', () => session.findText({ text: 'A', within: { kind: 'story', story: 'body' }, view: 'accepted' }));
  check('readStories', () => {
    const result = session.readStories({ view: 'accepted' });
    if (!result.ok) throw new Error(JSON.stringify(result.failure));
  });
  check('listStories', () => session.listStories());
  check('exportStructured', () => session.exportStructured({ revisionView: 'accepted' }));
  check('exportMarkdown', () => session.exportMarkdown({ revisionView: 'accepted' }));
  check('listComments', () => session.listComments());
  check('listRevisions', () => session.listRevisions());
  for (const story of danglingCells(session)) {
    const loc: YrsLoc = { story, paraId: `${story}#p0`, offset: 0 };
    check(`${story} paragraphs`, () => expectOneEmptyParagraph(story, session.paragraphs(story)));
    check(`${story} paragraphSpans`, () => {
      const spans = session.paragraphSpans(story);
      if (spans.length !== 1 || spans[0]!.paraId !== loc.paraId) throw new Error(JSON.stringify(spans));
    });
    check(`${story} hasStory`, () => { if (!session.hasStory(story)) throw new Error('hasStory is false'); });
    check(`${story} storySegments`, () => session.storySegments(story));
    check(`${story} storySegmentUnits`, () => session.storySegmentUnits(story, [0]));
    check(`${story} storySegmentUnitDigests`, () => session.storySegmentUnitDigests(story));
    check(`${story} storyLength`, () => session.storyLength(story));
    check(`${story} storyChecksum`, () => session.storyChecksum(story));
    check(`${story} yrsBlocksForStory`, () => session.yrsBlocksForStory(story));
    check(`${story} locateParagraph`, () => session.locateParagraph(story, loc.paraId));
    check(`${story} readParagraphs`, () => {
      const result = session.readParagraphs({ story, view: 'accepted' });
      if (!result.ok) throw new Error(JSON.stringify(result.failure));
      expectOneEmptyParagraph(story, result.paragraphs);
    });
    check(`${story} plainText`, () => { const text = storyPlainText(session, story); if (text !== '') throw new Error(JSON.stringify(text)); });
    check(`${story} selectionContext`, () => session.selectionContext(collapsed(loc)));
    check(`${story} selectionText`, () => session.selectionText(collapsed(loc)));
    check(`${story} projection`, () => {
      const projection = createYrsPositionProjection(session, 'body');
      if (!projection) throw new Error('no body projection');
      if (!projection.allTables().some((table) => table.cells.some((entry) => entry.story === story))) throw new Error('no projected cell');
      if (projection.positionForLoc(loc) === null) throw new Error('no projected position');
    });
    check(`${story} resolveLiveCell`, () => { if (!resolveLiveCell(commandReader(session), story)) throw new Error('no live cell'); });
    check(`${story} boundaryDelete`, () => {
      for (const direction of ['backward', 'forward'] as const) {
        tableBoundaryDeleteTarget({
          paragraphSpans: (id) => session.paragraphSpans(id),
          paragraphSegments: (id, indices) => session.storySegmentUnits(id, indices),
        }, loc, direction);
      }
    });
    check(`${story} sticky`, () => {
      const resolved = session.resolveStickyPosition(session.encodeStickyPosition(loc));
      if (resolved?.story !== story || resolved.offset !== 0) throw new Error(JSON.stringify(resolved));
    });
    check(`${story} selection`, () => {
      session.setSelection(loc, loc);
      const current = session.selection();
      if (current?.anchor.story !== story || current.head.story !== story) throw new Error(JSON.stringify(current));
      const cursor = session.encodeSelection();
      if (cursor && session.resolveSelection(cursor)?.anchor.story !== story) throw new Error('cursor does not resolve into the cell');
    });
  }
  return rows;
}

/** Every reader succeeds and reading writes nothing. */
function expectReadable(session: YrsSession): void {
  const before = stateBytes(session);
  for (const [reader, outcome] of Object.entries(readerRows(session))) expect(`${reader}: ${outcome}`).toBe(`${reader}: ok`);
  expect(stateBytes(session) === before).toBe(true);
}

async function expectSavedSound(session: YrsSession, clientId: number): Promise<void> {
  const saved = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
  const reopened = createSession(clientId);
  reopened.openDocx(saved, true);
  reopened.registerFont(font());
  expect(danglingCells(reopened)).toEqual([]);
  expectReadable(reopened);
}

for (const order of ['synced', 'concurrent'] as const) {
  for (const count of [2, 3] as const) {
    test(`${count} peers, ${order}: a named-but-missing cell reads as one empty paragraph on every replica, a late joiner and a state-only open`, async () => {
      const base = 9800 + (order === 'synced' ? 0 : 50) + count * 10;
      const { a, peers, story } = await danglingRoom(base, count, order);
      const joiner = lateJoiner(a, base + 4);
      const opened = stateOnly(a, base + 5);
      console.log(`READERS ${order} ${count} ${JSON.stringify({ story, missing: danglingCells(a) })}`);
      for (const replica of [...peers, joiner, opened]) {
        expect(danglingCells(replica)).toContain(story);
        expectReadable(replica);
        expect(resolveLiveCell(commandReader(replica), story)).toMatchObject({ story: 'body', tableIndex: 0, row: cellOf(replica, story).row });
      }
      expectConverged([...peers, joiner]);
      for (const [index, replica] of [...peers, joiner].entries()) await expectSavedSound(replica, base + 6 + index);
    }, 60_000);
  }
}

type WriteThunk = () => unknown;

interface WriteCase {
  name: string;
  /** Sets up the write (it may write elsewhere first) and returns the write aimed at the cell. */
  prepare: (session: YrsSession, story: string) => WriteThunk;
  /** A collapsed deletion at a cell edge may be a no-op instead of a refusal. */
  mayNoop?: boolean;
}

const cellLoc = (session: YrsSession, story: string) => firstParagraph(session, story);
const cellRange = (session: YrsSession, story: string) => collapsed(cellLoc(session, story));

/** Every write entry point a user, the editor or a host can aim at one missing cell. */
const WRITES: WriteCase[] = [
  { name: 'insertText', prepare: (s, story) => () => s.insertText(cellLoc(s, story), 'typed') },
  { name: 'insertText suggesting (tracked insertion)', prepare: (s, story) => () => s.insertText(cellLoc(s, story), 'typed', AUTHOR) },
  {
    name: 'applyInput (resident typing)',
    prepare: (s, story) => {
      s.layoutDocumentWithRegionsRetainedJson(REGION_REQUEST);
      s.buildDisplayListFrame('{}', 0);
      const loc = cellLoc(s, story);
      s.setSelection(loc, loc);
      return () => s.applyInput('x', s.residentCaretSnapshot().frameEpoch);
    },
  },
  {
    name: 'applyDelete backward (resident Backspace)', mayNoop: true,
    prepare: (s, story) => {
      s.layoutDocumentWithRegionsRetainedJson(REGION_REQUEST);
      s.buildDisplayListFrame('{}', 0);
      const loc = cellLoc(s, story);
      s.setSelection(loc, loc);
      return () => s.applyDelete('backward', s.residentCaretSnapshot().frameEpoch);
    },
  },
  {
    name: 'applyDelete forward (resident Delete)', mayNoop: true,
    prepare: (s, story) => {
      s.layoutDocumentWithRegionsRetainedJson(REGION_REQUEST);
      s.buildDisplayListFrame('{}', 0);
      const loc = cellLoc(s, story);
      s.setSelection(loc, loc);
      return () => s.applyDelete('forward', s.residentCaretSnapshot().frameEpoch);
    },
  },
  { name: 'replaceRange (composition commit, paste)', prepare: (s, story) => () => s.replaceRange(cellRange(s, story), 'pasted') },
  { name: 'deleteRange collapsed', mayNoop: true, prepare: (s, story) => () => s.deleteRange(cellRange(s, story)) },
  { name: 'deleteRange suggesting (tracked deletion)', mayNoop: true, prepare: (s, story) => () => s.deleteRange(cellRange(s, story), AUTHOR) },
  { name: 'clearFormatting', prepare: (s, story) => () => s.clearFormatting(cellRange(s, story)) },
  { name: 'mergeParagraphs', prepare: (s, story) => () => s.mergeParagraphs(story, cellLoc(s, story).paraId) },
  { name: 'toggleMark', prepare: (s, story) => () => s.toggleMark(cellRange(s, story), { type: 'bold' }) },
  { name: 'formatRange', prepare: (s, story) => () => s.formatRange(cellRange(s, story), { bold: true }) },
  { name: 'applyParagraphStyle', prepare: (s, story) => () => s.applyParagraphStyle(cellRange(s, story), 'Heading1') },
  { name: 'setParagraphAttrs', prepare: (s, story) => () => s.setParagraphAttrs(cellRange(s, story), { alignment: 'center' }) },
  { name: 'setHyperlink', prepare: (s, story) => () => s.setHyperlink(cellRange(s, story), { href: 'https://example.com' }) },
  { name: 'insertLineBreak', prepare: (s, story) => () => s.insertLineBreak(cellLoc(s, story)) },
  { name: 'splitParagraph', prepare: (s, story) => () => s.splitParagraph(cellLoc(s, story)) },
  { name: 'insertTable', prepare: (s, story) => () => s.insertTable(cellLoc(s, story), 1, 1) },
  { name: 'insertImage', prepare: (s, story) => () => s.insertImage(cellLoc(s, story), { rId: 'rIdImage' }) },
  { name: 'insertPageBreak', prepare: (s, story) => () => s.insertPageBreak(cellLoc(s, story)) },
  { name: 'insertSectionBreak', prepare: (s, story) => () => s.insertSectionBreak(cellLoc(s, story), 'nextPage') },
  {
    name: 'insertWatermark',
    prepare: (s, story) => () => s.insertWatermark(cellLoc(s, story), {
      kind: 'text', text: 'DRAFT', font: 'Liberation Sans', color: 'C0C0C0', semitransparent: true, layout: 'diagonal',
    }),
  },
  {
    name: 'moveEmbed destination',
    prepare: (s, story) => {
      const receipt = s.insertImage(firstParagraph(s, 'body'), { rId: 'rIdImage' });
      const from = receipt.range!.start;
      return () => natives.get(s)!.move_embed('body', from.paraId, from.offset, story, cellLoc(s, story).paraId, 0);
    },
  },
  { name: 'addComment', prepare: (s, story) => () => s.addComment([cellRange(s, story)], AUTHOR.name, AUTHOR.date, 'note') },
  {
    name: 'setCommentRanges onto the cell',
    prepare: (s, story) => {
      const anchor = firstParagraph(s, 'body');
      const { commentId } = s.addComment([{ story: 'body', start: { paraId: anchor.paraId, offset: 0 }, end: { paraId: anchor.paraId, offset: 6 } }], AUTHOR.name, AUTHOR.date, 'note');
      return () => s.setCommentRanges(commentId, [cellRange(s, story)]);
    },
  },
  { name: 'applyRawOps', prepare: (s, story) => () => s.applyRawOps(story, [{ op: 'insert', index: 0, text: 'raw' }]) },
  {
    name: 'mergeCells with a neighbour',
    prepare: (s, story) => () => { const target = cellOf(s, story); return s.mergeCells({ anchor: { ...target, column: Math.max(0, target.column - 1) }, head: target }); },
  },
  { name: 'splitCell', prepare: (s, story) => () => s.splitCell(cellOf(s, story), 1, 2) },
  {
    name: 'host batch insertText step',
    prepare: (s, story) => () => s.applyEdits({
      expectVersion: s.version(),
      steps: [{ op: 'insertText', target: { kind: 'paragraph', story, paraId: cellLoc(s, story).paraId }, at: 'end', text: 'from host' }],
    }),
  },
  {
    name: 'host batch replaceText step',
    prepare: (s, story) => () => s.applyEdits({
      expectVersion: s.version(),
      steps: [{ op: 'replaceText', target: { kind: 'paragraph', story, paraId: cellLoc(s, story).paraId }, text: 'from host' }],
    }),
  },
  {
    name: 'host batch insertParagraphs step',
    prepare: (s, story) => () => s.applyEdits({
      expectVersion: s.version(),
      steps: [{ op: 'insertParagraphs', target: { story, paraId: cellLoc(s, story).paraId }, at: 'end', paragraphs: [{ text: 'from host' }] }],
    }),
  },
  {
    name: 'host batch setParagraphStyle step',
    prepare: (s, story) => () => s.applyEdits({
      expectVersion: s.version(),
      steps: [{ op: 'setParagraphStyle', target: { story, paraId: cellLoc(s, story).paraId }, styleId: 'Heading1' }],
    }),
  },
];

function batchUnsupported(value: unknown): boolean {
  const result = value as { ok?: boolean; failure?: { code?: string } } | null | undefined;
  return !!result && typeof result === 'object' && result.ok === false && result.failure?.code === 'unsupported';
}

test('every write entry point aimed at a missing cell refuses before any mutation, and the room stays readable and converged', async () => {
  const outcomes: Record<string, string> = {};
  const problems: string[] = [];
  let base = 10_000;
  for (const { name, prepare, mayNoop } of WRITES) {
    const { a, b, peers, story } = await danglingRoom(base, 2, 'synced');
    base += 10;
    const run = prepare(b, story);
    syncAll(peers);
    const before = peers.map(stateBytes);
    let error: unknown;
    let value: unknown;
    try { value = run(); } catch (caught) { error = caught; }
    syncAll(peers);
    const unchanged = peers.every((peer, index) => stateBytes(peer) === before[index]);
    const text = error === undefined ? `ok ${JSON.stringify(value ?? null).slice(0, 120)}` : `throws: ${fullMessage(error)}`;
    outcomes[name] = `${text.slice(0, 200)}${unchanged ? '' : ' (wrote)'}`;
    if (/unreachable|panicked|RuntimeError/.test(text)) problems.push(`${name}: engine trap ${text}`);
    const refused = (error !== undefined && fullMessage(error).includes('has no story')) || batchUnsupported(value);
    if (!refused && !(mayNoop && error === undefined)) problems.push(`${name}: not refused: ${text.slice(0, 200)}`);
    if (!unchanged) problems.push(`${name}: state changed`);
    for (const peer of [a, b]) {
      for (const [reader, result] of Object.entries(readerRows(peer))) if (result !== 'ok') problems.push(`${name}: ${reader} ${result}`);
    }
    if (content(a) !== content(b)) problems.push(`${name}: peers diverge`);
    if (!danglingCells(b).includes(story)) problems.push(`${name}: the cell was recreated`);
  }
  console.log(`WRITE-SURFACE ${JSON.stringify(outcomes)}`);
  console.log(`WRITE-SURFACE-PROBLEMS ${JSON.stringify(problems)}`);
  expect(problems).toEqual([]);
}, 300_000);

test('writes elsewhere in the same table and document apply and converge while the cell stays missing', async () => {
  const { a, b, peers, story } = await danglingRoom(10_400, 3, 'synced');
  const named = ((b.tablePayload('body', 0)?.rows as { cells?: { story?: string }[] }[] | undefined) ?? [])
    .flatMap((row) => (row.cells ?? []).map((entry) => entry.story ?? ''));
  const other = named.find((id) => id !== '' && id !== story && b.storyIds().includes(id));
  expect(other).toBeDefined();
  b.insertText(firstParagraph(b, other!), 'typed elsewhere');
  peers[2]!.insertText(firstParagraph(peers[2]!, 'body'), 'Body ');
  a.insertRow(cellOf(a, other!), 'below');
  syncAll(peers);
  for (const peer of peers) {
    expect(cellText(peer, other!)).toContain('typed elsewhere');
    expect(cellText(peer, 'body')).toContain('Body ');
    expect(danglingCells(peer)).toContain(story);
    expectReadable(peer);
  }
  expectConverged(peers);
}, 60_000);

const RECOVERIES: Array<[string, (room: { a: YrsSession; b: YrsSession; story: string }) => unknown]> = [
  ['A redo', ({ a }) => expect(a.redo()).toBe(true)],
  ['deleting the row that names the cell', ({ b, story }) => { const target = cellOf(b, story); return b.deleteRow({ anchor: target, head: target }); }],
  ['deleting the table', ({ b }) => b.deleteTable({ story: 'body', tableIndex: 0 })],
];

for (const [index, [name, recover]] of RECOVERIES.entries()) {
  test(`${name} recovers: no missing cell, every reader works and peers converge`, async () => {
    const base = 10_500 + index * 20;
    const room = await danglingRoom(base, 3, 'synced');
    recover(room);
    syncAll(room.peers);
    for (const peer of room.peers) {
      expect(danglingCells(peer)).toEqual([]);
      expectReadable(peer);
    }
    expectConverged(room.peers);
    await expectSavedSound(room.a, base + 9);
  }, 60_000);
}

test('a nested table copied by a merge reads as empty cells after its insertion is undone', async () => {
  const a = await open(10_601);
  const b = await peerOf(a, 10_602);
  for (const peer of [a, b]) peer.registerFont(font());
  const cells = (a.tablePayload('body', 0) as { rows: { cells: { story: string }[] }[] }).rows[0]!.cells;
  const right = cells[1]!.story;
  a.insertTable(firstParagraph(a, right), 1, 2);
  sync(a, b);
  b.mergeCells({ anchor: at(0, 0), head: at(0, 1) });
  sync(a, b);
  expect(a.undo()).toBe(true);
  sync(a, b);
  const missing = { a: danglingCells(a), b: danglingCells(b) };
  console.log(`MERGED-NESTED ${JSON.stringify(missing)}`);
  expect(missing.a.length).toBeGreaterThan(0);
  for (const peer of [a, b, lateJoiner(a, 10_603)]) {
    for (const story of danglingCells(peer)) expect(peer.paragraphs(story).map((entry) => entry.text)).toEqual(['']);
    expectReadable(peer);
  }
  expectConverged([a, b]);
}, 60_000);

test('placing a caret in a missing cell and reading it writes nothing; a restored caret resolves without writing', async () => {
  const { b, story } = await danglingRoom(9940, 2, 'synced');
  const before = stateBytes(b);
  const vector = Array.from(b.encodeStateVector());
  const caret = firstParagraph(b, story);
  expect(caret.paraId).toBe(`${story}#p0`);
  b.setSelection(caret, caret);
  expect(b.selection()).toMatchObject({ anchor: { story, offset: 0 }, head: { story, offset: 0 } });
  const sticky = b.encodeStickyPosition(caret);
  expect(b.resolveStickyPosition(sticky)).toMatchObject({ story, offset: 0 });
  const cursor = b.encodeSelection();
  expectReadable(b);
  expect(stateBytes(b) === before).toBe(true);

  const restored = await peerOf(b, 9949);
  restored.registerFont(font());
  const restoredBefore = stateBytes(restored);
  restored.setSelection(caret, caret);
  expect(restored.selection()).toMatchObject({ anchor: { story, offset: 0 } });
  expect(restored.resolveStickyPosition(sticky)).toMatchObject({ story, offset: 0 });
  if (cursor) expect(restored.resolveSelection(cursor)).toMatchObject({ anchor: { story, offset: 0 } });
  expectReadable(restored);
  expect(stateBytes(restored) === restoredBefore).toBe(true);
  expect(Array.from(restored.encodeStateVector())).toEqual(vector);
});

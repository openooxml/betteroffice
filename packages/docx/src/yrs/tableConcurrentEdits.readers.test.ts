import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { createYrsSidebarProjection } from '../layout/render/yrsSidebarProjection';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { createEditSession, preloadEditWasm, type EditSession } from '../wasm/edit';
import { saveYrsDocx, type YrsAuthor, type YrsCellBorders, type YrsCellLoc, type YrsLoc, type YrsSession, type YrsStoryRange, type YrsTableRange } from './index';
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

/** Layout, the display frame and the body projection all succeed on the session. */
function expectRenders(session: YrsSession): void {
  const outcomes: string[] = [];
  const check = (name: string, run: () => unknown) => {
    try { run(); } catch (error) { outcomes.push(`${name} throws: ${message(error)}`); }
  };
  check('layout', () => session.layoutDocumentWithRegionsRetainedJson(REGION_REQUEST));
  check('frame', () => session.buildDisplayListFrame('{}', 0));
  check('projection', () => createYrsPositionProjection(session, 'body')?.allTables());
  expect(outcomes).toEqual([]);
}

test('undoing a split deletes the text another collaborator typed into the new cell, and both peers still render', async () => {
  const a = await open(9751);
  const b = await peerOf(a, 9752);
  for (const peer of [a, b]) peer.registerFont(font());
  a.splitCell(at(0, 1), 1, 2);
  sync(a, b);
  const created = a.storyIds().filter((story) => story.includes('~9751'));
  expect(created.length).toBeGreaterThan(0);
  const target = created[0]!;
  const [para] = b.paragraphs(target);
  b.insertText({ story: target, paraId: para!.paraId, offset: 0 }, 'typed by B');
  sync(a, b);
  expect({ a: cellText(a, target), b: cellText(b, target) }).toEqual({ a: 'typed by B', b: 'typed by B' });
  expect(a.undo()).toBe(true);
  sync(a, b);
  for (const peer of [a, b]) {
    expect(peer.storyIds()).not.toContain(target);
    expect(peer.storyIds().some((story) => (cellText(peer, story) ?? '').includes('typed by B'))).toBe(false);
    expect(danglingCells(peer)).toEqual([]);
    expectRenders(peer);
  }
  expectConverged([a, b]);
});

const FURTHER_STEPS: Array<[string, (room: { a: YrsSession; b: YrsSession; peers: YrsSession[] }) => unknown]> = [
  ['A undo again', ({ a }) => expect(a.undo()).toBe(false)],
  ['B undo', ({ b }) => expect(b.undo()).toBe(true)],
  ['B inserts another row', ({ b }) => b.insertRow(at(0, 0), 'below')],
  ['C inserts a row below the new one', ({ peers }) => peers[2]!.insertRow(at(1, 0), 'below')],
];

for (const order of ['synced', 'concurrent'] as const) {
  for (const [index, [name, step]] of FURTHER_STEPS.entries()) {
    test(`${order}: ${name} after the cell went missing applies, and every peer and a late joiner render and read`, async () => {
      const base = 9770 + (order === 'synced' ? 0 : 200) + index * 10;
      const room = await danglingRoom(base, 3, order);
      step(room);
      syncAll(room.peers);
      const joiner = lateJoiner(room.a, base + 9);
      for (const peer of [...room.peers, joiner]) {
        expectRenders(peer);
        expectReadable(peer);
      }
      expectConverged([...room.peers, joiner]);
    }, 60_000);
  }
}

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
    check(`${story} plainText`, () => { const text = storyPlainText(session, story); if (text !== '\n') throw new Error(JSON.stringify(text)); });
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
  /** A host-side argument check that rejects every table cell before the engine is reached. */
  hostRefusal?: RegExp;
}

const cellLoc = (session: YrsSession, story: string) => firstParagraph(session, story);
const cellRange = (session: YrsSession, story: string) => collapsed(cellLoc(session, story));

const CELL_TEXT_FORMAT = { margins: { left: 240, right: 108 } };
const CELL_BORDERS: YrsCellBorders = { top: { style: 'single', size: 4, color: { rgb: '000000' } } };
const only = (session: YrsSession, story: string): YrsTableRange => { const target = cellOf(session, story); return { anchor: target, head: target }; };
const fromFirst = (session: YrsSession, story: string): YrsTableRange => ({ anchor: at(0, 0), head: cellOf(session, story) });

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
  { name: 'insertSectionBreak', hostRefusal: /^Section breaks belong to body paragraphs$/, prepare: (s, story) => () => s.insertSectionBreak(cellLoc(s, story), 'nextPage') },
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
  { name: 'createStory of the missing id', prepare: (s, story) => () => s.createStory(story, 'revived') },
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

/** The engine's refusal of a write into a cell with no story: an `Error` whose `code` is `cell-has-no-story`. */
function cellRefused(error: unknown): boolean {
  return error instanceof Error && 'code' in error && error.code === 'cell-has-no-story';
}

function batchUnsupported(value: unknown): boolean {
  const result = value as { ok?: boolean; failure?: { code?: string; reason?: string } } | null | undefined;
  return !!result && typeof result === 'object' && result.ok === false && result.failure?.code === 'unsupported' &&
    result.failure.reason === 'cell-has-no-story';
}

for (const order of ['synced', 'concurrent'] as const) test(`${order}: every write entry point aimed at a missing cell refuses before any mutation, and the room stays readable and converged`, async () => {
  const outcomes: Record<string, string> = {};
  const problems: string[] = [];
  let base = order === 'synced' ? 10_000 : 30_000;
  for (const { name, prepare, mayNoop, hostRefusal } of WRITES) {
    const { a, b, peers, story } = await danglingRoom(base, 2, order);
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
    const refused = cellRefused(error) || batchUnsupported(value) || (hostRefusal !== undefined && error instanceof Error && hostRefusal.test(error.message));
    if (!refused && !(mayNoop && error === undefined)) problems.push(`${name}: not refused: ${text.slice(0, 200)}`);
    if (!unchanged) problems.push(`${name}: state changed`);
    for (const peer of [a, b]) {
      for (const [reader, result] of Object.entries(readerRows(peer))) if (result !== 'ok') problems.push(`${name}: ${reader} ${result}`);
    }
    if (content(a) !== content(b)) problems.push(`${name}: peers diverge`);
    if (!danglingCells(b).includes(story)) problems.push(`${name}: the cell was recreated`);
  }
  console.log(`WRITE-SURFACE ${order} ${JSON.stringify(outcomes)}`);
  console.log(`WRITE-SURFACE-PROBLEMS ${order} ${JSON.stringify(problems)}`);
  expect(problems).toEqual([]);
}, 300_000);

for (const order of ['synced', 'concurrent'] as const) test(`${order}: writes elsewhere in the same table and document apply and converge while the cell stays missing`, async () => {
  const { a, b, peers, story } = await danglingRoom(order === 'synced' ? 10_400 : 32_000, 3, order);
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

/** Whether `setCellTextFormat` writes the cell's story (refused on a missing cell) or only the table payload (allowed). */
const CELL_TEXT_FORMAT_REFUSED = false;

const otherCell = (): YrsTableRange => ({ anchor: at(0, 0), head: at(0, 0) });

type CellProperties = Record<string, unknown> | undefined;

/** Name, whether it is refused on the missing cell, the write, and whether its change shows in a cell's `tcPr`. */
const CELL_PROPERTY_WRITES: Array<[string, boolean, (s: YrsSession, range: YrsTableRange) => unknown, (tcPr: CellProperties) => boolean]> = [
  ['setCellShading', false, (s, range) => s.setCellShading(range, 'FF0000'), (tcPr) => tcPr?.backgroundColor === 'FF0000'],
  [
    'setCellBorders', false, (s, range) => s.setCellBorders(range, CELL_BORDERS),
    (tcPr) => (tcPr?.borders as { top?: { style?: string } } | undefined)?.top?.style === 'single',
  ],
  [
    'setCellTextFormat', CELL_TEXT_FORMAT_REFUSED, (s, range) => s.setCellTextFormat(range, CELL_TEXT_FORMAT),
    (tcPr) => JSON.stringify(tcPr?.margins ?? null).includes('240'),
  ],
];

/** The `tcPr` the body table's payload holds for the cell named `story`. */
function cellProperties(session: YrsSession, story: string): CellProperties {
  const payload = session.tablePayload('body', 0) as { rows?: { cells?: { story?: string; tcPr?: Record<string, unknown> }[] }[] } | null;
  for (const row of payload?.rows ?? []) {
    const found = (row.cells ?? []).find((entry) => entry.story === story);
    if (found) return found.tcPr;
  }
  throw new Error(`no cell names ${story}`);
}

for (const order of ['synced', 'concurrent'] as const) {
  test(`${order}: cell property writes on the missing cell, alone and in a range, follow the contract and apply on other cells`, async () => {
    let base = order === 'synced' ? 10_800 : 33_000;
    for (const [name, refused, write, landed] of CELL_PROPERTY_WRITES) {
      for (const scope of ['the cell', 'a range including it', 'another cell'] as const) {
        const room = await danglingRoom(base, 2, order);
        base += 10;
        const first = tableCellIds(room.b)[0]!;
        expect(first).not.toBe(room.story);
        for (const peer of room.peers) expect([landed(cellProperties(peer, room.story)), landed(cellProperties(peer, first))]).toEqual([false, false]);
        const range = scope === 'the cell' ? only(room.b, room.story) : scope === 'another cell' ? otherCell() : fromFirst(room.b, room.story);
        const before = room.peers.map(stateBytes);
        let error: unknown;
        try { write(room.b, range); } catch (caught) { error = caught; }
        syncAll(room.peers);
        const label = `${name} on ${scope}`;
        const outcome = `${label}: ${error === undefined ? 'applied' : fullMessage(error)}`;
        const expectRefusal = refused && scope !== 'another cell';
        if (expectRefusal) {
          expect(`${label}: ${cellRefused(error) ? 'cell-has-no-story' : outcome}`).toBe(`${label}: cell-has-no-story`);
          expect(room.peers.every((peer, index) => stateBytes(peer) === before[index])).toBe(true);
        } else {
          expect(outcome).toBe(`${label}: applied`);
        }
        const onMissing = !expectRefusal && scope !== 'another cell';
        const onFirst = !expectRefusal && scope !== 'the cell';
        for (const peer of room.peers) {
          expect({ label, missing: landed(cellProperties(peer, room.story)), first: landed(cellProperties(peer, first)) })
            .toEqual({ label, missing: onMissing, first: onFirst });
          expect(danglingCells(peer)).toContain(room.story);
          expectReadable(peer);
        }
        expectConverged(room.peers);
        if (!expectRefusal) await expectSavedSound(room.a, base - 1);
      }
    }
  }, 300_000);
}

function tableCellIds(session: YrsSession): string[] {
  return ((session.tablePayload('body', 0)?.rows as { cells?: { story?: string }[] }[] | undefined) ?? [])
    .flatMap((row) => (row.cells ?? []).map((entry) => entry.story ?? ''));
}

const TABLE_STEPS: Array<[string, (s: YrsSession) => unknown]> = [
  ['split another cell into two columns', (s) => s.splitCell(at(0, 0), 1, 2)],
  ['split a cell of the inserted row', (s) => s.splitCell(at(1, 0), 1, 2)],
  ['merge the two halves of that cell', (s) => s.mergeCells({ anchor: at(1, 0), head: at(1, 1) })],
  ['insert a row', (s) => s.insertRow(at(1, 0), 'below')],
  ['insert a column', (s) => s.insertColumn(at(1, 0), 'right')],
];

for (const order of ['synced', 'concurrent'] as const) {
  test(`${order}: splitting another cell, merging, and inserting rows and columns keep working without recreating the missing id`, async () => {
    const base = order === 'synced' ? 10_900 : 34_000;
    const { b, peers, story } = await danglingRoom(base, 3, order);
    for (const [name, step] of TABLE_STEPS) {
      let error: unknown;
      try { step(b); } catch (caught) { error = caught; }
      expect(`${name}: ${error === undefined ? 'applied' : fullMessage(error)}`).toBe(`${name}: applied`);
      syncAll(peers);
      for (const peer of peers) {
        expect(peer.storyIds()).not.toContain(story);
        expect(danglingCells(peer)).toContain(story);
        const ids = tableCellIds(peer);
        expect(new Set(ids).size).toBe(ids.length);
        expectReadable(peer);
      }
      expectConverged(peers);
    }
  }, 120_000);

  test(`${order}: a caret placed in the new cell before it went missing reads at the start of its empty paragraph on every peer`, async () => {
    const base = order === 'synced' ? 10_950 : 34_100;
    const a = await open(base + 1);
    const peers = [a, await peerOf(a, base + 2), await peerOf(a, base + 3)];
    for (const peer of peers) peer.registerFont(font());
    const b = peers[1]!;
    a.splitCell(at(0, 1), 1, 2);
    syncAll(peers);
    const target = a.storyIds().find((story) => story.includes(`~${base + 1}`));
    expect(target).toBeDefined();
    for (const peer of peers) {
      const loc = firstParagraph(peer, target!);
      peer.setSelection(loc, loc);
    }
    const saved = b.encodeSelection();
    expect(saved).not.toBeNull();
    b.insertRow(at(0, 0), 'below');
    if (order === 'synced') syncAll(peers);
    expect(a.undo()).toBe(true);
    syncAll(peers);
    const caret = { story: target!, paraId: `${target!}#p0`, offset: 0 };
    for (const peer of peers) {
      expect(danglingCells(peer)).toContain(target!);
      const before = stateBytes(peer);
      let selection: ReturnType<YrsSession['selection']> = null;
      expect(() => { selection = peer.selection(); }).not.toThrow();
      expect(selection).toMatchObject({ anchor: caret, head: caret });
      expect(() => peer.selectionContext(collapsed(caret))).not.toThrow();
      expect(() => peer.encodeSelection()).not.toThrow();
      expect(stateBytes(peer) === before).toBe(true);
      expectReadable(peer);
    }
    expectConverged(peers);
    for (const replica of [peers[2]!, lateJoiner(a, base + 4)]) {
      const before = stateBytes(replica);
      let restored: ReturnType<YrsSession['resolveSelection']> = null;
      expect(() => { restored = replica.resolveSelection(saved!); }).not.toThrow();
      expect(restored).toMatchObject({ anchor: caret, head: caret });
      expect(stateBytes(replica) === before).toBe(true);
    }
  }, 60_000);

  test(`${order}: loadStories with a fresh story and then the missing cell is refused as a whole`, async () => {
    const { b, story } = await danglingRoom(order === 'synced' ? 10_990 : 34_200, 2, order);
    const fresh = 'fresh-story';
    const before = stateBytes(b);
    let error: unknown;
    try {
      b.loadStories([
        { storyId: fresh, paragraphs: [{ text: 'fresh' }] },
        { storyId: story, paragraphs: [{ text: 'revived' }] },
      ]);
    } catch (caught) { error = caught; }
    expect(cellRefused(error) ? 'cell-has-no-story' : `${error === undefined ? 'loaded' : fullMessage(error)}`).toBe('cell-has-no-story');
    expect(stateBytes(b) === before).toBe(true);
    expect(b.storyIds()).not.toContain(fresh);
    expect(danglingCells(b)).toContain(story);
    expectReadable(b);
  }, 60_000);
}

const RECOVERIES: Array<[string, (room: { a: YrsSession; b: YrsSession; story: string }) => unknown]> = [
  ['A redo', ({ a }) => expect(a.redo()).toBe(true)],
  ['deleting the row that names the cell', ({ b, story }) => { const target = cellOf(b, story); return b.deleteRow({ anchor: target, head: target }); }],
  ['deleting the table', ({ b }) => b.deleteTable({ story: 'body', tableIndex: 0 })],
];

for (const order of ['synced', 'concurrent'] as const) for (const [index, [name, recover]] of RECOVERIES.entries()) {
  test(`${order}: ${name} recovers: no missing cell, every reader works and peers converge`, async () => {
    const base = (order === 'synced' ? 10_500 : 31_000) + index * 20;
    const room = await danglingRoom(base, 3, order);
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

for (const order of ['synced', 'concurrent'] as const) {
  test(`${order}: a comment anchored in the cell before it went missing resolves without throwing or writing`, async () => {
    const base = order === 'synced' ? 10_700 : 10_710;
    const a = await open(base + 1);
    const b = await peerOf(a, base + 2);
    for (const peer of [a, b]) peer.registerFont(font());
    a.splitCell(at(0, 1), 1, 2);
    sync(a, b);
    const target = a.storyIds().find((story) => story.includes(`~${base + 1}`));
    expect(target).toBeDefined();
    const paraId = b.paragraphSpans(target!)[0]!.paraId;
    b.insertText({ story: target!, paraId, offset: 0 }, 'noted');
    const { commentId } = b.addComment([{ story: target!, start: { paraId, offset: 0 }, end: { paraId, offset: 5 } }], AUTHOR.name, AUTHOR.date, 'note');
    if (order === 'synced') sync(a, b);
    b.insertRow(at(0, 0), 'below');
    if (order === 'synced') sync(a, b);
    expect(a.undo()).toBe(true);
    sync(a, b);
    for (const peer of [a, b, lateJoiner(a, base + 3)]) {
      expect(danglingCells(peer)).toContain(target!);
      const before = stateBytes(peer);
      expect(() => peer.resolveComment(commentId)).not.toThrow();
      expect(() => peer.listComments()).not.toThrow();
      expectReadable(peer);
      expect(stateBytes(peer) === before).toBe(true);
    }
    expectConverged([a, b]);
  }, 60_000);
}

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

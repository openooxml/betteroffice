import { afterEach, expect, spyOn, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { preflightDocxCommand } from './residentCommandActions';
import { hyperlinkAt, projectDocxEditState } from './residentCommandEditState';
import type { DocxCommandApplyResult, DocxWorkerCommandEngine } from './residentCommandProtocol';
import type { YrsHyperlinkAttrs, YrsLoc } from './index';
import type { EditSession } from './wasm/index';
import type { ResidentEngineSession } from './residentEngineSession';
import {
  TIMEOUT, closeCommandEngines, command, context, execute, expectCaret,
  openCommandEngine, selection, snapshot, texts,
} from './residentCommandTestSupport';

afterEach(closeCommandEngines);

test('d1b_apply_results_carry_selection_and_editstate', async () => {
  const arm = await openCommandEngine(['abc'], { editState: true });
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const complete = (result: DocxCommandApplyResult) => {
    expect(result).toHaveProperty('selection');
    expect(result).toHaveProperty('editState');
    expect(result.selection).toEqual(arm.state.selection);
    expect(result.editState).toEqual(projectDocxEditState(arm.engine, arm.state));
  };
  const applied = execute(arm, command(arm, 'insertText', { text: 'X' }));
  expect(applied.status).toBe('applied');
  expectCaret(applied, { story: 'body', paraId: id, offset: 1 });
  complete(applied);
  expect(applied.editState).toMatchObject({ readOnly: false, canUndo: true, canRedo: false });
  expect(texts(arm)).toEqual(['Xabc']);
  expect(arm.version).toBe(1);

  const beforeNoop = snapshot(arm);
  const noop = execute(arm, command(arm, 'insertText', { text: '' }));
  expect(noop.status).toBe('noop');
  complete(noop);
  expect(snapshot(arm)).toEqual(beforeNoop);
  const moved = execute(arm, command(arm, 'select', { selection: selection(id, 0, 2) }));
  expect(moved.status).toBe('noop');
  complete(moved);
  expect(moved.editState?.toolbar?.hasSelection).toBe(true);
  expect(arm.version).toBe(1);

  const invalid = command(arm, 'insertText', { text: 'blocked', selection: selection(id, 100) });
  const beforeRefusal = snapshot(arm);
  const refused = preflightDocxCommand(arm.engine, arm.state, invalid, context(arm, invalid));
  expect(refused.status).toBe('refused');
  if (refused.status !== 'refused') throw new Error('Invalid selection passed preflight');
  complete(refused);
  expect(snapshot(arm)).toEqual(beforeRefusal);

  const mode = execute(arm, command(arm, 'setReadOnly', { readOnly: true }));
  expect(mode.status).toBe('noop');
  complete(mode);
  const beforeReadOnlyRefusal = snapshot(arm);
  const readOnlyRefusal = execute(arm, command(arm, 'insertText', { text: 'blocked' }));
  expect(readOnlyRefusal.status).toBe('refused');
  complete(readOnlyRefusal);
  expect(readOnlyRefusal.editState?.readOnly).toBe(true);
  expect(snapshot(arm)).toEqual(beforeReadOnlyRefusal);
}, TIMEOUT);

test('d1b_toolbar_storedmarks_hyperlink_and_table_context_are_worker_reads', async () => {
  const bodyXml =
    '<w:p w14:paraId="00000001"><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Bold</w:t></w:r>' +
    '<w:r><w:t xml:space="preserve"> plain </w:t></w:r>' +
    '<w:hyperlink r:id="rIdLink" w:tooltip="D1 link"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>' +
    '<w:tbl><w:tblPr><w:jc w:val="center"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:shd w:fill="AABBCC"/></w:tcPr><w:p w14:paraId="00000002"><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p w14:paraId="00000003"><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const arm = await openCommandEngine([], { editState: true, bodyXml });
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const reads = spyOn(arm.engine, 'selectionContext');
  const segments = spyOn(arm.engine, 'paragraphSegments');
  const cells = spyOn(arm.engine, 'cellSelection');
  const tables = spyOn(arm.engine, 'tablePayload');
  try {
    const bold = execute(arm, command(arm, 'select', { selection: selection(id, 0, 4) }));
    expect(bold.status).toBe('noop');
    const boldState = projectDocxEditState(arm.engine, arm.state);
    expect(boldState.toolbar).toMatchObject({ bold: true, fontSize: 28, alignment: 'center' });
    expect(boldState.effectiveFormatting).toMatchObject({ bold: true, fontSizePt: 14, alignment: 'center' });
    expect(boldState.storedMarks).toBeNull();
    expect(boldState.hyperlink).toBeNull();
    expect(boldState.tableContext).toBeNull();
    expect(execute(arm, command(arm, 'select', { selection: selection(id, 0, 10) })).status).toBe('noop');
    expect(projectDocxEditState(arm.engine, arm.state).toolbar?.bold).toBe('mixed');

    expect(execute(arm, command(arm, 'select', { selection: selection(id, 6) })).status).toBe('noop');
    const beforeMarks = snapshot(arm);
    expect(execute(arm, command(arm, 'toggleMark', { mark: { type: 'italic' } })).status).toBe('noop');
    const marked = projectDocxEditState(arm.engine, arm.state);
    expect(marked.storedMarks).toMatchObject({ story: 'body', paraId: id, delta: { italic: true } });
    expect(marked.effectiveFormatting?.italic).toBe(true);
    expect(snapshot(arm)).toMatchObject({ texts: beforeMarks.texts, ids: beforeMarks.ids, vector: beforeMarks.vector, version: 0, canUndo: false, canRedo: false });

    expect(execute(arm, command(arm, 'select', { selection: selection(id, 13) })).status).toBe('noop');
    expect(projectDocxEditState(arm.engine, arm.state).hyperlink).toMatchObject({ href: 'https://example.com/d1', tooltip: 'D1 link' });
    expect(execute(arm, command(arm, 'select', { selection: selection(id, 6) })).status).toBe('noop');
    expect(projectDocxEditState(arm.engine, arm.state).hyperlink).toBeNull();

    const cellStories = arm.engine.geometryReader.storyIds().filter((story) =>
      arm.engine.paragraphs(story).some(({ text }) => text === 'Cell A' || text === 'Cell B')
    );
    expect(cellStories).toHaveLength(2);
    const cellStory = cellStories.find((story) => arm.engine.paragraphs(story)[0]?.text === 'Cell A');
    if (!cellStory) throw new Error('Synthetic table cell story was not seeded');
    const cellId = arm.engine.paragraphs(cellStory)[0]!.paraId;
    expect(execute(arm, command(arm, 'select', { selection: selection(cellId, 1, 1, cellStory) })).status).toBe('noop');
    const cellState = projectDocxEditState(arm.engine, arm.state);
    expect(cellState.toolbar?.inTable).toBe(true);
    expect(cellState.tableContext).toMatchObject({
      story: 'body', rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 2,
      hasMultiCellSelection: false, canSplitCell: true, justification: 'center', cellBackgroundColor: 'AABBCC',
    });
    const tableIndex = cellState.tableContext?.tableIndex;
    if (tableIndex === undefined) throw new Error('Worker table context lacks index');
    arm.native.set_cell_selection(JSON.stringify({
      anchor: { story: 'body', tableIndex, row: 0, column: 0 },
      head: { story: 'body', tableIndex, row: 0, column: 1 },
    }));
    expect(projectDocxEditState(arm.engine, arm.state).tableContext).toMatchObject({
      rowIndex: 0, columnIndex: 1, hasMultiCellSelection: true, rowCount: 1, columnCount: 2,
    });
    expect(reads).toHaveBeenCalled();
    expect(segments).toHaveBeenCalled();
    expect(cells).toHaveBeenCalled();
    expect(tables).toHaveBeenCalled();
    expect(arm.version).toBe(0);
    expect(arm.native.can_undo()).toBe(false);
    const readOnly = execute(arm, command(arm, 'setReadOnly', { readOnly: true }));
    expect(readOnly.status).toBe('noop');
    expect(readOnly.editState?.readOnly).toBe(true);
  } finally {
    tables.mockRestore(); cells.mockRestore(); segments.mockRestore(); reads.mockRestore();
  }
}, TIMEOUT);

test('worker_table_properties_state_matches_toolbar_shape', async () => {
  const bodyXml =
    '<w:p w14:paraId="00000001"><w:r><w:t>Before</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:jc w:val="right"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4" w:color="112233"/></w:tcBorders><w:shd w:fill="AABBCC"/></w:tcPr>' +
    '<w:p w14:paraId="00000002"><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4" w:color="445566"/></w:tcBorders></w:tcPr>' +
    '<w:p w14:paraId="00000003"><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const arm = await openCommandEngine([], { editState: true, bodyXml });
  const storyA = arm.engine.geometryReader.storyIds().find((story) => arm.engine.paragraphs(story)[0]?.text === 'Cell A')!;
  const idA = arm.engine.paragraphs(storyA)[0]!.paraId;
  expect(execute(arm, command(arm, 'select', { selection: selection(idA, 0, 0, storyA) })).status).toBe('noop');
  arm.engine.setCellSelection({ anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 }, head: { story: 'body', tableIndex: 0, row: 0, column: 1 } });
  const state = projectDocxEditState(arm.engine, arm.state);
  expect(state.tableContext).toMatchObject({
    rowIndex: 0, columnIndex: 1, justification: 'right',
    cellBorderColor: { rgb: '112233' }, cellBackgroundColor: 'AABBCC',
  });
}, TIMEOUT);

test('d1b_undo_redo_availability_excludes_noops_refusals_and_remote_updates', async () => {
  const arm = await openCommandEngine(['abc'], { editState: true, clientId: 8201 });
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const initial = arm.engine.encodeState();
  arm.engine.beginUndoCapture();
  const availability = (canUndo: boolean, canRedo: boolean) => {
    expect(projectDocxEditState(arm.engine, arm.state)).toMatchObject({ canUndo, canRedo });
    expect(arm.native.can_undo()).toBe(canUndo);
    expect(arm.native.can_redo()).toBe(canRedo);
  };
  availability(false, false);
  for (const intent of [
    command(arm, 'insertText', { text: '' }),
    command(arm, 'deleteBackward', { selection: selection(id, 0) }),
    command(arm, 'select', { selection: selection(id, 0) }),
    command(arm, 'toggleMark', { mark: { type: 'bold' } }),
  ]) {
    expect(execute(arm, intent).status).toBe('noop');
    availability(false, false);
    expect(arm.version).toBe(0);
  }
  const beforeRefusal = snapshot(arm);
  expect(execute(arm, command(arm, 'insertText', { text: 'blocked', selection: selection(id, 999) })).status).toBe('refused');
  expect(snapshot(arm)).toEqual(beforeRefusal);
  availability(false, false);
  expect(execute(arm, command(arm, 'insertText', { text: 'X' })).status).toBe('applied');
  availability(true, false);
  expect(arm.version).toBe(1);
  expect(arm.native.undo()).toBe(true);
  expect(texts(arm)).toEqual(['abc']);
  availability(false, true);

  expect(execute(arm, command(arm, 'select', { selection: selection(id, 0) })).status).toBe('noop');
  expect(execute(arm, command(arm, 'insertText', { text: '' })).status).toBe('noop');
  const beforeRedoRefusal = snapshot(arm);
  expect(execute(arm, command(arm, 'paragraphBreak', { selection: selection(id, -1) })).status).toBe('refused');
  expect(snapshot(arm)).toEqual(beforeRedoRefusal);
  availability(false, true);

  const remote = await openCommandEngine(['abc'], { editState: true, clientId: 8202, checkpoint: initial });
  const vector = arm.engine.encodeStateVector();
  const remoteId = remote.engine.paragraphs('body')[0]!.paraId;
  remote.engine.insertText({ story: 'body', paraId: remoteId, offset: 3 }, 'R');
  arm.engine.applyUpdate(remote.engine.encodeStateAsUpdate(vector));
  expect(texts(arm)).toEqual(['abcR']);
  expect(arm.version).toBe(1);
  availability(false, true);
  expect(arm.native.redo()).toBe(true);
  expect(texts(arm)).toEqual(['XabcR']);
  availability(true, false);
  expect(arm.native.undo()).toBe(true);
  availability(false, true);
  expect(texts(arm)).toEqual(['abcR']);
  expect(execute(arm, command(arm, 'select', { selection: selection(id, 0) })).status).toBe('noop');
  expect(execute(arm, command(arm, 'insertText', { text: 'Y' })).status).toBe('applied');
  availability(true, false);
  expect(texts(arm)).toEqual(['YabcR']);
  expect(arm.version).toBe(2);
  expect(arm.native.undo()).toBe(true);
  expect(texts(arm)).toEqual(['abcR']);
  availability(false, true);
}, TIMEOUT);

const CORPUS = ['crates/docx-edit/tests/fixtures', 'crates/betteroffice-docx/tests/corpus/fixtures', 'packages/docx/src/yrs/__fixtures__'];
const AUTHOR = ['Caret', '2026-01-01T00:00:00Z'] as const;

function corpusDocuments(): string[] {
  const found: string[] = [];
  const visit = (path: string) => {
    if (statSync(path).isDirectory()) for (const entry of readdirSync(path).sort()) visit(join(path, entry));
    else if (path.endsWith('.docx')) found.push(path);
  };
  for (const root of CORPUS) visit(resolve(process.cwd(), root));
  return found;
}

function seeded(seed: string): () => number {
  let state = 2166136261;
  for (const character of seed) state = Math.imul(state ^ character.charCodeAt(0), 16777619);
  return () => {
    state = (state + 0x6d2b79f5) | 0;
    let value = Math.imul(state ^ (state >>> 15), 1 | state);
    value = (value + Math.imul(value ^ (value >>> 7), 61 | value)) ^ value;
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

type CaretEngine = Pick<ResidentEngineSession, 'geometryReader' | 'paragraphSegments' | 'paragraphUnit' | 'revisionIdsAt'>;

function storyScanHyperlinkAt(engine: CaretEngine, story: string, paraId: string, offset: number): YrsHyperlinkAttrs | null {
  const index = engine.geometryReader.paragraphSpans(story).findIndex((paragraph) => paragraph.paraId === paraId);
  if (index < 0) return null;
  let start = 0;
  let boundary: YrsHyperlinkAttrs | null = null;
  for (const segment of engine.paragraphSegments(story, [index])[0] ?? []) {
    const end = start + (segment.kind === 'text' ? segment.text.length : 1);
    const link = segment.attributes.hyperlink as YrsHyperlinkAttrs | undefined;
    if (link && typeof link.href === 'string' && link.href) {
      const attrs = {
        href: link.href,
        ...(typeof link.tooltip === 'string' ? { tooltip: link.tooltip } : {}),
        ...(typeof link.rId === 'string' ? { rId: link.rId } : {}),
      };
      if (offset > start && offset < end) return attrs;
      if (!boundary && (offset === start || offset === end)) boundary = attrs;
    }
    start = end;
  }
  return boundary;
}

function documentScanRevisionsAt(engine: CaretEngine, head: YrsLoc): string[] {
  const reader = engine.geometryReader;
  const at = (paraId: string, offset: number) => reader.locateParagraph(head.story, paraId).start + offset;
  const position = at(head.paraId, head.offset);
  return reader.listRevisions().filter(({ range }) => {
    if (range.story !== head.story) return false;
    try {
      return at(range.start.paraId, range.start.offset) <= position && position <= at(range.end.paraId, range.end.offset);
    } catch {
      return false;
    }
  }).map(({ revisionId }) => revisionId);
}

/** Carets at every revision end and hyperlink edge (and their neighbours), each paragraph's ends, and random offsets. */
function carets(engine: CaretEngine, random: () => number): YrsLoc[] {
  const reader = engine.geometryReader;
  const found: YrsLoc[] = [];
  const spans = new Map(reader.storyIds().map((story) => [story, reader.paragraphSpans(story)]));
  const near = (story: string, paraId: string, offset: number) => {
    const span = spans.get(story)?.find((entry) => entry.paraId === paraId);
    if (!span) return;
    for (const at of [offset - 1, offset, offset + 1]) if (at >= 0 && at <= span.length) found.push({ story, paraId, offset: at });
  };
  for (const { range } of reader.listRevisions()) {
    near(range.story, range.start.paraId, range.start.offset);
    near(range.story, range.end.paraId, range.end.offset);
  }
  for (const [story, list] of spans) {
    list.forEach((span, index) => {
      found.push({ story, paraId: span.paraId, offset: 0 }, { story, paraId: span.paraId, offset: span.length },
        { story, paraId: span.paraId, offset: Math.floor(random() * (span.length + 1)) });
      if (index > 400) return;
      let start = 0;
      for (const segment of engine.paragraphSegments(story, [index])[0] ?? []) {
        const end = start + (segment.kind === 'text' ? segment.text.length : 1);
        if (segment.attributes.hyperlink) { near(story, span.paraId, start); near(story, span.paraId, end); }
        start = end;
      }
    });
  }
  return found.length > 3000 ? found.filter(() => random() < 3000 / found.length) : found;
}

function suggestEdit(engine: CaretEngine, native: EditSession, random: () => number): boolean {
  const reader = engine.geometryReader;
  const stories = reader.storyIds();
  const story = random() < 0.6 && stories.includes('body') ? 'body' : stories[Math.floor(random() * stories.length)]!;
  const spans = reader.paragraphSpans(story);
  if (!spans.length) return false;
  const index = Math.floor(random() * spans.length);
  const span = spans[index]!;
  const offset = Math.floor(random() * (span.length + 1));
  const next = spans[Math.min(index + 1 + Math.floor(random() * 2), spans.length - 1)]!;
  const kind = Math.floor(random() * 4);
  try {
    if (kind === 0) native.insert_text(story, span.paraId, offset, 'caret', ...AUTHOR);
    else if (kind === 1) native.split_paragraph(story, span.paraId, offset, ...AUTHOR);
    else if (kind === 2) native.delete_range(story, span.paraId, offset, next.paraId, next === span ? span.length : Math.floor(random() * (next.length + 1)), ...AUTHOR);
    else native.insert_text(story, span.paraId, offset, 'plain');
    return true;
  } catch {
    return false;
  }
}

function expectCaretReadsMatch(engine: CaretEngine, label: string, random: () => number, seen: { carets: number; links: number; revisions: number; noteCarets: number }) {
  for (const at of carets(engine, random)) {
    const where = `${label} ${at.story} ${at.paraId}@${at.offset}`;
    const link = hyperlinkAt(engine, at.story, at.paraId, at.offset);
    expect(link, where).toEqual(storyScanHyperlinkAt(engine, at.story, at.paraId, at.offset));
    const revisions = engine.revisionIdsAt(at);
    expect(revisions, where).toEqual(documentScanRevisionsAt(engine, at));
    seen.carets++;
    if (link) seen.links++;
    if (revisions.length) seen.revisions++;
    if (/^(fn|en):/.test(at.story)) seen.noteCarets++;
  }
}

const SYNTHETIC_REVISIONS = (() => {
  const stamp = (id: number) => `w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"`;
  return '<w:p w14:paraId="0000000A"><w:pPr><w:rPr><w:ins ' + stamp(11) + '/></w:rPr></w:pPr>' +
    '<w:r><w:t xml:space="preserve">Lead </w:t></w:r><w:ins ' + stamp(11) + '><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>' +
    '<w:p w14:paraId="0000000B"><w:ins ' + stamp(11) + '><w:r><w:t>still inserted</w:t></w:r></w:ins>' +
    '<w:r><w:t xml:space="preserve"> then </w:t></w:r><w:hyperlink r:id="rIdLink" w:tooltip="T"><w:r><w:t>link</w:t></w:r>' +
    '<w:ins ' + stamp(12) + '><w:r><w:t>ed</w:t></w:r></w:ins></w:hyperlink>' +
    '<w:del ' + stamp(13) + '><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>' +
    '<w:p w14:paraId="0000000B"><w:hyperlink r:id="rIdLink"><w:r><w:t>Dup</w:t></w:r></w:hyperlink></w:p>' +
    '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="0000000C">' +
    '<w:hyperlink r:id="rIdLink"><w:r><w:t>cell link</w:t></w:r></w:hyperlink>' +
    '<w:ins ' + stamp(14) + '><w:r><w:t xml:space="preserve"> cell ins</w:t></w:r></w:ins></w:p></w:tc></w:tr></w:tbl>' +
    '<w:p w14:paraId="0000000D"><w:r><w:t>Tail</w:t></w:r></w:p>';
})();

/** Unique paragraph ids: a mark revision before the caret paragraph, an insertion across a formatting boundary, a row insertion. */
const SYNTHETIC_UNIQUE_IDS = (() => {
  const stamp = (id: number) => `w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"`;
  return '<w:p w14:paraId="0000001A"><w:pPr><w:rPr><w:ins ' + stamp(21) + '/></w:rPr></w:pPr>' +
    '<w:r><w:t xml:space="preserve">Lead </w:t></w:r><w:ins ' + stamp(21) + '><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>' +
    '<w:p w14:paraId="0000001B"><w:ins ' + stamp(21) + '><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">still </w:t></w:r>' +
    '<w:r><w:t>inserted</w:t></w:r></w:ins><w:del ' + stamp(23) + '><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>' +
    '<w:p w14:paraId="0000001C"><w:pPr><w:rPr><w:del ' + stamp(24) + '/></w:rPr></w:pPr><w:r><w:t>Mark deleted</w:t></w:r></w:p>' +
    '<w:p w14:paraId="0000001D"><w:r><w:t>After</w:t></w:r></w:p>' +
    '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:ins ' + stamp(25) + '/></w:trPr><w:tc>' +
    '<w:p w14:paraId="0000001E"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    '<w:p w14:paraId="0000001F"><w:r><w:t>Tail</w:t></w:r></w:p>';
})();

test('caret hyperlink and revision reads match whole-document scans over the corpus and suggested edits', async () => {
  const seen = { carets: 0, links: 0, revisions: 0, noteCarets: 0 };
  for (const [label, bodyXml] of [['synthetic', SYNTHETIC_REVISIONS], ['synthetic unique ids', SYNTHETIC_UNIQUE_IDS]] as const) {
    const synthetic = await openCommandEngine([], { editState: true, bodyXml });
    const syntheticRandom = seeded(label);
    expectCaretReadsMatch(synthetic.engine, label, syntheticRandom, seen);
    for (let edit = 0; edit < 12; edit++) {
      if (suggestEdit(synthetic.engine, synthetic.native, syntheticRandom)) expectCaretReadsMatch(synthetic.engine, `${label} edit ${edit}`, syntheticRandom, seen);
    }
  }
  const unique = await openCommandEngine([], { editState: true, bodyXml: SYNTHETIC_UNIQUE_IDS });
  const at = (paraId: string, offset: number) => unique.engine.revisionIdsAt({ story: 'body', paraId, offset }).length;
  expect([at('0000001B', 0), at('0000001B', 3), at('0000001D', 0), at('0000001D', 1)]).toEqual([2, 1, 1, 0]);

  const wasm = await import('./wasm/index');
  const { createResidentEngineSession } = await import('./residentEngineSession');
  let opened = 0;
  for (const path of corpusDocuments()) {
    const label = relative(process.cwd(), path);
    const create = wasm.createEditSession;
    let native: EditSession | undefined;
    const capture = spyOn(wasm, 'createEditSession').mockImplementation((clientId) => (native = create(clientId)));
    let engine: ResidentEngineSession;
    try { engine = await createResidentEngineSession(undefined, 8301); } finally { capture.mockRestore(); }
    try {
      try { engine.openDocx(new Uint8Array(readFileSync(path)), undefined, 'caret-local'); } catch { continue; }
      opened++;
      const random = seeded(label);
      expectCaretReadsMatch(engine, label, random, seen);
      for (let edit = 0; edit < 3; edit++) {
        if (suggestEdit(engine, native!, random)) expectCaretReadsMatch(engine, `${label} edit ${edit}`, random, seen);
      }
    } finally {
      engine.destroy();
    }
  }
  expect(opened).toBeGreaterThan(20);
  expect(seen.links).toBeGreaterThan(5);
  expect(seen.revisions).toBeGreaterThan(20);
  expect(seen.noteCarets).toBeGreaterThan(0);
  console.log(`caret-local parity: ${opened} documents, ${JSON.stringify(seen)}`);
}, 600_000);

test('edit state reads the caret paragraph alone, not every paragraph or revision', async () => {
  const arm = await openCommandEngine([], { editState: true, bodyXml: SYNTHETIC_REVISIONS });
  const engine: DocxWorkerCommandEngine = arm.engine;
  const linked = engine.paragraphs('body')[1]!;
  expect(linked.text.slice(0, 5)).toBe('still');
  expect(linked.text.slice(20, 24)).toBe('link');
  const spans = spyOn(engine.geometryReader, 'paragraphSpans');
  const revisions = spyOn(engine.geometryReader, 'listRevisions');
  try {
    expect(execute(arm, command(arm, 'select', { selection: selection(linked.paraId, 22) })).status).toBe('noop');
    spans.mockClear(); revisions.mockClear();
    expect(projectDocxEditState(arm.engine, arm.state).hyperlink)
      .toMatchObject({ href: 'https://example.com/d1', tooltip: 'T' });
    expect(spans).not.toHaveBeenCalled();
    expect(revisions).not.toHaveBeenCalled();
    expect(execute(arm, command(arm, 'select', { selection: selection(linked.paraId, 3) })).status).toBe('noop');
    spans.mockClear(); revisions.mockClear();
    expect(projectDocxEditState(arm.engine, arm.state).currentRevisionId).toBeDefined();
    expect(spans).not.toHaveBeenCalled();
    expect(revisions).not.toHaveBeenCalled();
  } finally {
    spans.mockRestore(); revisions.mockRestore();
  }
}, TIMEOUT);

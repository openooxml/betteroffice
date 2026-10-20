import { afterEach, expect, spyOn, test } from 'bun:test';
import { preflightDocxCommand } from './residentCommandActions';
import { projectDocxEditState } from './residentCommandEditState';
import type { DocxCommandApplyResult } from './residentCommandProtocol';
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

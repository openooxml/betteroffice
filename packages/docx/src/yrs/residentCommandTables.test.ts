import { afterEach, expect, test } from 'bun:test';
import {
  TIMEOUT, closeCommandEngines, execute, openCommandEngine, selection,
  type CommandArm,
} from './residentCommandTestSupport';
import type { DocxCommandEnvelope } from './residentCommandProtocol';
import { tableAnchors, type TablePayload } from './storyPlainText';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, properties = '') => `<w:tc><w:tcPr>${properties}</w:tcPr>${paragraph(text)}</w:tc>`;
const table = (rows: string[]) => `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
const body = paragraph('Before') + table([cell('A') + cell('B'), cell('C') + cell('D')]) + paragraph('After');

type Arm = Awaited<ReturnType<typeof openCommandEngine>>;

function tableCommand(arm: CommandArm, kind: string, payload: Record<string, unknown>): DocxCommandEnvelope {
  arm.seq += 1;
  return {
    session: 'd1-worker', incarnation: 1, seq: arm.seq,
    basis: { epoch: 1, version: arm.version }, kind, payload,
  } as unknown as DocxCommandEnvelope;
}

function payload(arm: Arm, story = 'body', index = 0): TablePayload | null {
  return arm.engine.tablePayload(story, index) as unknown as TablePayload | null;
}

function anchors(arm: Arm, story = 'body', index = 0) {
  return tableAnchors(payload(arm, story, index)!).anchors;
}

function cellStory(arm: Arm, text: string): string {
  const story = arm.engine.geometryReader.storyIds()
    .find((story) => story.includes(':t') && arm.engine.paragraphs(story)[0]?.text === text);
  if (!story) throw new Error(`Cell ${text} missing`);
  return story;
}

function cellText(arm: Arm, story: string): string {
  return arm.engine.paragraphs(story).map((paragraph) => paragraph.text).join('\n');
}

async function inCell(text = 'A', bodyXml = body): Promise<Arm> {
  const arm = await openCommandEngine([], { bodyXml });
  const story = cellStory(arm, text);
  const paraId = arm.engine.paragraphs(story)[0]!.paraId;
  expect(execute(arm, tableCommand(arm, 'select', { selection: selection(paraId, 0, 0, story) })).status).toBe('noop');
  return arm;
}

async function inCellOn(arm: Arm, text: string): Promise<void> {
  const story = cellStory(arm, text);
  const paraId = arm.engine.paragraphs(story)[0]!.paraId;
  expect(execute(arm, tableCommand(arm, 'select', { selection: selection(paraId, 0, 0, story) })).status).toBe('noop');
}

function validFocus(arm: Arm, result: { selection?: unknown }): { story: string; paraId: string; offset: number } {
  const focus = (result.selection as { focus: { story: string; paraId: string; offset: number } } | null)?.focus;
  expect(focus).toBeDefined();
  expect(arm.engine.geometryReader.hasStory(focus!.story)).toBe(true);
  expect(arm.engine.geometryReader.paragraphSpans(focus!.story).some((span) => span.paraId === focus!.paraId)).toBe(true);
  return focus!;
}

function selectCells(arm: Arm, anchor: [number, number], head: [number, number], index = 0): void {
  const at = ([row, column]: [number, number]) => ({ story: 'body', tableIndex: index, row, column });
  arm.engine.setCellSelection({ anchor: at(anchor), head: at(head) });
}

function typedAfterTable(arm: Arm, text: string, ordinal = 0): boolean {
  const segments = arm.engine.geometryReader.storySegments('body');
  const tables = segments.flatMap((segment, index) => segment.kind === 'embed' && segment.embedKind === 'table' ? [index] : []);
  const typed = segments.findIndex((segment) => segment.kind === 'text' && segment.text.includes(text));
  const next = tables[ordinal + 1] ?? Infinity;
  return tables[ordinal] !== undefined && tables[ordinal]! < typed && typed < next;
}

function undo(arm: Arm) {
  return execute(arm, tableCommand(arm, 'undo', {}));
}

function redo(arm: Arm) {
  return execute(arm, tableCommand(arm, 'redo', {}));
}

afterEach(closeCommandEngines);

test('insert_table_row_below_applies_undoes_redoes_and_types', async () => {
  const arm = await inCell();
  const before = JSON.stringify(payload(arm));
  const story = cellStory(arm, 'A');
  const applied = execute(arm, tableCommand(arm, 'insertTableRow', { side: 'below' }));
  expect(applied.status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(3);
  expect(validFocus(arm, applied).story).toBe(story);
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  validFocus(arm, { selection: arm.state.selection });
  expect(redo(arm).status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(3);
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'x' })).status).toBe('applied');
  expect(cellText(arm, story)).toBe('xA');
}, TIMEOUT);

test('insert_table_row_above_targets_the_focused_row', async () => {
  const arm = await inCell('C');
  const before = JSON.stringify(payload(arm));
  const story = cellStory(arm, 'C');
  const applied = execute(arm, tableCommand(arm, 'insertTableRow', { side: 'above' }));
  expect(applied.status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(3);
  expect(cellText(arm, story)).toBe('C');
  expect(validFocus(arm, applied).story).toBe(story);
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 2, column: 0 }, head: { row: 2, column: 0 } });
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(redo(arm).status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(3);
}, TIMEOUT);

test('insert_table_column_right_applies_undoes_redoes_and_types', async () => {
  const arm = await inCell();
  const before = JSON.stringify(payload(arm));
  const story = cellStory(arm, 'A');
  const applied = execute(arm, tableCommand(arm, 'insertTableColumn', { side: 'right' }));
  expect(applied.status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(3);
  expect(validFocus(arm, applied).story).toBe(story);
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(redo(arm).status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(3);
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'x' })).status).toBe('applied');
  expect(cellText(arm, story)).toBe('xA');
}, TIMEOUT);

test('insert_table_column_left_targets_the_focused_column', async () => {
  const arm = await inCell('B');
  const before = JSON.stringify(payload(arm));
  const story = cellStory(arm, 'B');
  const applied = execute(arm, tableCommand(arm, 'insertTableColumn', { side: 'left' }));
  expect(applied.status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(3);
  expect(validFocus(arm, applied).story).toBe(story);
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 2 }, head: { row: 0, column: 2 } });
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(redo(arm).status).toBe('applied');
}, TIMEOUT);

test('delete_table_row_removes_the_focused_row_and_restores_on_undo', async () => {
  const arm = await inCell('C');
  const before = JSON.stringify(payload(arm));
  const survivor = cellStory(arm, 'A');
  const applied = execute(arm, tableCommand(arm, 'deleteTableRow', {}));
  expect(applied.status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(1);
  expect(validFocus(arm, applied).story).toBe(survivor);
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(cellText(arm, cellStory(arm, 'C'))).toBe('C');
  expect(redo(arm).status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(1);
  const after = execute(arm, tableCommand(arm, 'insertText', { text: 'x' }));
  expect(after.status).toBe('applied');
  expect(cellText(arm, validFocus(arm, after).story).startsWith('x')).toBe(true);
}, TIMEOUT);

test('delete_table_column_removes_the_focused_column_and_restores_on_undo', async () => {
  const arm = await inCell('B');
  const before = JSON.stringify(payload(arm));
  const survivor = cellStory(arm, 'A');
  const applied = execute(arm, tableCommand(arm, 'deleteTableColumn', {}));
  expect(applied.status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(1);
  expect(validFocus(arm, applied).story).toBe(survivor);
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(redo(arm).status).toBe('applied');
}, TIMEOUT);

test('delete_table_column_from_the_last_column_keeps_the_caret_in_its_row', async () => {
  const arm = await inCell('D');
  const survivor = cellStory(arm, 'C');
  const applied = execute(arm, tableCommand(arm, 'deleteTableColumn', {}));
  expect(applied.status).toBe('applied');
  expect(tableAnchors(payload(arm)!).columns).toBe(1);
  expect(validFocus(arm, applied).story).toBe(survivor);
}, TIMEOUT);

test('delete_table_removes_the_table_and_types_in_the_following_paragraph', async () => {
  const arm = await inCell();
  const applied = execute(arm, tableCommand(arm, 'deleteTable', {}));
  expect(applied.status).toBe('applied');
  expect(payload(arm)).toBeNull();
  const caret = validFocus(arm, applied);
  expect(caret.story).toBe('body');
  expect(undo(arm).status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(2);
  expect(cellText(arm, cellStory(arm, 'A'))).toBe('A');
  expect(redo(arm).status).toBe('applied');
  expect(payload(arm)).toBeNull();
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'T' })).status).toBe('applied');
  expect(arm.engine.paragraphs('body').map((paragraph) => paragraph.text)).toContain('TAfter');
}, TIMEOUT);

test('merge_table_cells_merges_split_restores_and_round_trips', async () => {
  const arm = await inCell();
  const before = JSON.stringify(payload(arm));
  const storyA = cellStory(arm, 'A');
  const paraA = arm.engine.paragraphs(storyA)[0]!.paraId;
  expect(execute(arm, tableCommand(arm, 'mergeTableCells', {
    selection: selection(paraA, 0, 0, storyA),
  }))).toMatchObject({ status: 'refused', reason: 'invalid-selection' });
  selectCells(arm, [0, 0], [0, 1]);
  const ranged = execute(arm, tableCommand(arm, 'mergeTableCells', {}));
  expect(ranged.status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(1);
  expect(anchors(arm)[0]).toMatchObject({ column: 0, colspan: 2 });
  expect(validFocus(arm, ranged).story).toBe(storyA);
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(redo(arm).status).toBe('applied');
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'x' })).status).toBe('applied');
  expect(cellText(arm, storyA).startsWith('x')).toBe(true);
  expect(execute(arm, tableCommand(arm, 'splitTableCell', { rows: 1, columns: 2 })).status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(2);
  expect(undo(arm).status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(1);
  expect(redo(arm).status).toBe('applied');
  expect(anchors(arm).filter((anchor) => anchor.row === 0)).toHaveLength(2);
  expect(undo(arm).status).toBe('applied');
  expect(undo(arm).status).toBe('applied');
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
}, TIMEOUT);

test('merge_table_cells_refuses_a_range_covering_only_one_merged_cell', async () => {
  const arm = await inCell();
  selectCells(arm, [0, 0], [0, 1]);
  expect(execute(arm, tableCommand(arm, 'mergeTableCells', {})).status).toBe('applied');
  expect(anchors(arm)[0]).toMatchObject({ column: 0, colspan: 2 });
  await inCellOn(arm, 'A');
  selectCells(arm, [0, 0], [0, 1]);
  const before = JSON.stringify(payload(arm));
  const version = arm.engine.proposalEngine.version();
  const canUndo = arm.engine.canUndo();
  expect(execute(arm, tableCommand(arm, 'mergeTableCells', {})))
    .toMatchObject({ status: 'refused', reason: 'invalid-selection' });
  expect(arm.engine.proposalEngine.version()).toBe(version);
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(arm.engine.canUndo()).toBe(canUndo);
}, TIMEOUT);

test('split_table_cell_subdivides_and_reports_size_refusals', async () => {
  const arm = await inCell();
  const storyA = cellStory(arm, 'A');
  const version = arm.engine.proposalEngine.version();
  expect(execute(arm, tableCommand(arm, 'splitTableCell', { rows: 1, columns: 1 })))
    .toMatchObject({ status: 'refused', reason: 'invalid-table-size' });
  expect(arm.engine.proposalEngine.version()).toBe(version);
  selectCells(arm, [0, 0], [1, 1]);
  const merged = execute(arm, tableCommand(arm, 'mergeTableCells', {}));
  expect(merged.status).toBe('applied');
  expect(anchors(arm)).toHaveLength(1);
  const mergedJson = JSON.stringify(payload(arm));
  expect(execute(arm, tableCommand(arm, 'splitTableCell', { rows: 1, columns: 2 })))
    .toMatchObject({ status: 'refused', reason: 'invalid-table-size' });
  const split = execute(arm, tableCommand(arm, 'splitTableCell', { rows: 2, columns: 2 }));
  expect(split.status).toBe('applied');
  expect(anchors(arm)).toHaveLength(4);
  expect(validFocus(arm, split).story).toBe(storyA);
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(mergedJson);
  expect(redo(arm).status).toBe('applied');
  expect(anchors(arm)).toHaveLength(4);
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'x' })).status).toBe('applied');
  expect(cellText(arm, storyA).startsWith('x')).toBe(true);
}, TIMEOUT);

test('set_table_cell_shading_shades_clears_and_keeps_selection', async () => {
  const arm = await inCell();
  const before = JSON.stringify(payload(arm));
  const kept = arm.state.selection!.focus;
  const shaded = execute(arm, tableCommand(arm, 'setTableCellShading', { color: 'FF0000' }));
  expect(shaded.status).toBe('applied');
  expect(payload(arm)!.rows[0]!.cells[0]!.tcPr).toMatchObject({ backgroundColor: 'FF0000' });
  expect(shaded.selection?.focus).toMatchObject({ story: kept.story, paraId: kept.paraId, offset: kept.offset });
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(redo(arm).status).toBe('applied');
  expect(payload(arm)!.rows[0]!.cells[0]!.tcPr).toMatchObject({ backgroundColor: 'FF0000' });
  const storyA = cellStory(arm, 'A');
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'x' })).status).toBe('applied');
  expect(cellText(arm, storyA)).toBe('xA');
  expect(execute(arm, tableCommand(arm, 'setTableCellShading', { color: null })).status).toBe('applied');
  expect(payload(arm)!.rows[0]!.cells[0]!.tcPr?.backgroundColor).toBeUndefined();
}, TIMEOUT);

test('table_commands_refuse_outside_cells_readonly_and_bad_payloads_without_mutation', async () => {
  const arm = await openCommandEngine([], { bodyXml: body });
  const kinds: [string, Record<string, unknown>][] = [
    ['insertTableRow', { side: 'below' }],
    ['insertTableColumn', { side: 'right' }],
    ['deleteTableRow', {}],
    ['deleteTableColumn', {}],
    ['deleteTable', {}],
    ['mergeTableCells', {}],
    ['splitTableCell', { rows: 1, columns: 2 }],
    ['setTableCellShading', { color: 'FF0000' }],
  ];
  for (const [kind, extra] of kinds) {
    const version = arm.engine.proposalEngine.version();
    const vector = arm.engine.encodeStateVector();
    expect(execute(arm, tableCommand(arm, kind, extra)))
      .toMatchObject({ status: 'refused', reason: 'invalid-cell-target' });
    expect(arm.engine.proposalEngine.version()).toBe(version);
    expect(arm.engine.encodeStateVector()).toEqual(vector);
  }
  await inCellOn(arm, 'A');
  for (const [kind, extra] of kinds) {
    arm.state.readOnly = true;
    const version = arm.engine.proposalEngine.version();
    expect(execute(arm, tableCommand(arm, kind, extra)))
      .toMatchObject({ status: 'refused', reason: 'read-only' });
    expect(arm.engine.proposalEngine.version()).toBe(version);
    arm.state.readOnly = false;
  }
  const bad: [string, Record<string, unknown>, string][] = [
    ['insertTableRow', { side: 'sideways' }, 'invalid-payload'],
    ['insertTableColumn', {}, 'invalid-payload'],
    ['splitTableCell', { rows: 0, columns: 2 }, 'invalid-table-size'],
    ['splitTableCell', { rows: 100, columns: 101 }, 'invalid-table-size'],
    ['setTableCellShading', { color: '#FF0000' }, 'invalid-payload'],
    ['setTableCellShading', { color: 'FFF' }, 'invalid-payload'],
    ['setTableCellShading', {}, 'invalid-payload'],
  ];
  for (const [kind, extra, reason] of bad) {
    const version = arm.engine.proposalEngine.version();
    expect(execute(arm, tableCommand(arm, kind, extra))).toMatchObject({ status: 'refused', reason });
    expect(arm.engine.proposalEngine.version()).toBe(version);
  }
  expect(execute(arm, tableCommand(arm, 'mergeTableCells', {})))
    .toMatchObject({ status: 'refused', reason: 'invalid-selection' });
  expect(execute(arm, tableCommand(arm, 'insertTableColumn', { side: 'right', suggesting: { name: 'W', date: '2026-01-01' } })))
    .toMatchObject({ status: 'refused', reason: 'unsupported-suggesting' });
  expect(execute(arm, tableCommand(arm, 'deleteTable', { suggesting: { name: 'W', date: '2026-01-01' } })))
    .toMatchObject({ status: 'refused', reason: 'unsupported-suggesting' });
}, TIMEOUT);

test('deleting_the_last_row_removes_the_table_beside_it', async () => {
  const xml = paragraph('Before') + table([cell('Only') + cell('Alone')]) + paragraph('After');
  const arm = await inCell('Only', xml);
  const applied = execute(arm, tableCommand(arm, 'deleteTableRow', {}));
  expect(applied.status).toBe('applied');
  expect(payload(arm)).toBeNull();
  const caret = validFocus(arm, applied);
  expect(caret.story).toBe('body');
  expect(arm.engine.paragraphs('body').find((paragraph) => paragraph.paraId === caret.paraId)?.text).toBe('After');
  expect(undo(arm).status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(1);
  expect(redo(arm).status).toBe('applied');
  expect(payload(arm)).toBeNull();
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'T' })).status).toBe('applied');
  expect(arm.engine.paragraphs('body').map((paragraph) => paragraph.text)).toContain('TAfter');
}, TIMEOUT);

test('nested_table_commands_target_the_inner_table', async () => {
  const xml = paragraph('Before') + table([
    cell('Outer') + `<w:tc>${paragraph('Host')}${table([cell('Inner A') + cell('Inner B')])}</w:tc>`,
  ]) + paragraph('After');
  const arm = await inCell('Inner A', xml);
  const host = cellStory(arm, 'Host');
  const outerBefore = JSON.stringify(payload(arm));
  const innerBefore = JSON.stringify(payload(arm, host, 0));
  expect(execute(arm, tableCommand(arm, 'insertTableColumn', { side: 'right' })).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(outerBefore);
  expect(tableAnchors(payload(arm, host, 0)!).columns).toBe(3);
  expect(undo(arm).status).toBe('applied');
  expect(JSON.stringify(payload(arm, host, 0))).toBe(innerBefore);
  expect(JSON.stringify(payload(arm))).toBe(outerBefore);
}, TIMEOUT);

test('explicit_selection_targets_that_cell_when_the_caret_is_elsewhere', async () => {
  const arm = await inCell('B');
  const storyC = cellStory(arm, 'C');
  const paraC = arm.engine.paragraphs(storyC)[0]!.paraId;
  const applied = execute(arm, tableCommand(arm, 'insertTableRow', {
    side: 'below', selection: selection(paraC, 0, 0, storyC),
  }));
  expect(applied.status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(3);
  expect(validFocus(arm, applied).story).toBe(storyC);
  expect(undo(arm).status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(2);
}, TIMEOUT);

test('cell_range_commands_apply_to_every_selected_cell', async () => {
  const arm = await inCell('A');
  selectCells(arm, [0, 0], [1, 1]);
  expect(execute(arm, tableCommand(arm, 'setTableCellShading', { color: '00FF00' })).status).toBe('applied');
  expect(payload(arm)!.rows.flatMap((row) => row.cells.map((cell) => cell.tcPr?.backgroundColor))).toEqual(['00FF00', '00FF00', '00FF00', '00FF00']);
  selectCells(arm, [0, 0], [1, 0]);
  const deleted = execute(arm, tableCommand(arm, 'deleteTableRow', {}));
  expect(deleted.status).toBe('applied');
  expect(payload(arm)).toBeNull();
  expect(validFocus(arm, deleted).story).toBe('body');
}, TIMEOUT);

test('undo_of_a_row_holding_the_caret_parks_it_after_the_table', async () => {
  const arm = await inCell('A');
  const before = JSON.stringify(payload(arm));
  expect(execute(arm, tableCommand(arm, 'insertTableRow', { side: 'below' })).status).toBe('applied');
  const added = anchors(arm).find((anchor) => anchor.row === 1 && anchor.column === 0)!.story;
  const paraId = arm.engine.paragraphs(added)[0]!.paraId;
  expect(execute(arm, tableCommand(arm, 'select', { selection: selection(paraId, 0, 0, added) })).status).toBe('noop');
  const undone = undo(arm);
  expect(undone.status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(validFocus(arm, undone).story).toBe('body');
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'x' })).status).toBe('applied');
  expect(JSON.stringify(payload(arm))).toBe(before);
  expect(typedAfterTable(arm, 'x')).toBe(true);
}, TIMEOUT);

test('undo_parks_the_caret_after_the_live_table_when_a_table_was_inserted_before_it', async () => {
  const arm = await openCommandEngine([], { bodyXml: body });
  const first = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, tableCommand(arm, 'select', { selection: selection(first, 6, 6, 'body') })).status).toBe('noop');
  expect(execute(arm, tableCommand(arm, 'insertTable', { rows: 1, columns: 1 })).status).toBe('applied');
  await inCellOn(arm, 'A');
  const tableIndex = arm.engine.geometryReader.storySegments('body')
    .filter((segment) => segment.kind === 'embed' && segment.embedKind === 'table').length - 1;
  expect(tableIndex).toBe(1);
  expect(execute(arm, tableCommand(arm, 'insertTableRow', { side: 'below' })).status).toBe('applied');
  const added = anchors(arm, 'body', 1).find((anchor) => anchor.row === 1 && anchor.column === 0)!.story;
  const paraId = arm.engine.paragraphs(added)[0]!.paraId;
  expect(execute(arm, tableCommand(arm, 'select', { selection: selection(paraId, 0, 0, added) })).status).toBe('noop');
  expect(undo(arm).status).toBe('applied');
  expect(payload(arm, 'body', 1)!.rows).toHaveLength(2);
  expect(execute(arm, tableCommand(arm, 'insertText', { text: 'x' })).status).toBe('applied');
  expect(typedAfterTable(arm, 'x', 1)).toBe(true);
}, TIMEOUT);

test('undo_in_a_nested_table_parks_the_caret_beside_the_inner_table', async () => {
  const xml = paragraph('Before') + table([
    cell('Outer') + `<w:tc>${paragraph('Host')}${table([cell('Inner A') + cell('Inner B')])}</w:tc>`,
  ]) + paragraph('After');
  const arm = await inCell('Inner A', xml);
  const host = cellStory(arm, 'Host');
  const innerBefore = JSON.stringify(payload(arm, host, 0));
  expect(execute(arm, tableCommand(arm, 'insertTableRow', { side: 'below' })).status).toBe('applied');
  const added = anchors(arm, host, 0).find((anchor) => anchor.row === 1 && anchor.column === 0)!.story;
  const paraId = arm.engine.paragraphs(added)[0]!.paraId;
  expect(execute(arm, tableCommand(arm, 'select', { selection: selection(paraId, 0, 0, added) })).status).toBe('noop');
  const undone = undo(arm);
  expect(undone.status).toBe('applied');
  expect(JSON.stringify(payload(arm, host, 0))).toBe(innerBefore);
  expect(validFocus(arm, undone).story).toBe(host);
}, TIMEOUT);

test('suggesting_insert_table_row_stamps_a_revision_that_reject_removes', async () => {
  const arm = await inCell();
  const author = { name: 'Writer', date: '2026-01-01T00:00:00Z' };
  const applied = execute(arm, tableCommand(arm, 'insertTableRow', { side: 'below', suggesting: author }));
  expect(applied.status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(3);
  const revisions = arm.engine.geometryReader.listRevisions() as { revisionId: string }[];
  expect(revisions.length).toBeGreaterThan(0);
  const rejected = execute(arm, tableCommand(arm, 'rejectRevision', {
    revisionId: revisions[0]!.revisionId, expectVersion: arm.engine.proposalEngine.version(),
  }));
  expect(rejected.status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(2);
}, TIMEOUT);

test('suggesting_delete_table_row_stamps_a_revision_that_reject_restores', async () => {
  const arm = await inCell('C');
  const author = { name: 'Writer', date: '2026-01-01T00:00:00Z' };
  expect(execute(arm, tableCommand(arm, 'deleteTableRow', { suggesting: author })).status).toBe('applied');
  const revisions = arm.engine.geometryReader.listRevisions() as { revisionId: string }[];
  expect(revisions.length).toBeGreaterThan(0);
  expect(execute(arm, tableCommand(arm, 'rejectRevision', {
    revisionId: revisions[0]!.revisionId, expectVersion: arm.engine.proposalEngine.version(),
  })).status).toBe('applied');
  expect(payload(arm)!.rows).toHaveLength(2);
  expect(cellText(arm, cellStory(arm, 'C'))).toBe('C');
}, TIMEOUT);

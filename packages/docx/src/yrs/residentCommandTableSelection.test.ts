import { afterEach, beforeAll, expect, test } from 'bun:test';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import type { TextRunPrimitive } from '../layout/render/displayList';
import { applyDocxCommand, preflightDocxCommand } from './residentCommandActions';
import { resolveShiftArrowCellRange } from './residentCommandTableSelection';
import { readResidentCellSelectionRects } from './residentObjectGeometry';
import type { DocxCommandEnvelope, DocxCommandSelection } from './residentCommandProtocol';
import { tableAnchors, type TablePayload } from './storyPlainText';
import { closeCommandEngines, command, context, execute, openCommandEngine as openEngine, selection, snapshot, TIMEOUT, type CommandArm } from './residentCommandTestSupport';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, properties = '') => `<w:tc><w:tcPr>${properties}</w:tcPr>${paragraph(text)}</w:tc>`;
const table = (rows: string[]) => `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
const body = paragraph('Before') + table([cell('A') + cell('B'), cell('C') + cell('D')]) + paragraph('After');

async function inCell(bodyXml = body, text = 'A') {
  const arm = await openEngine([], { bodyXml });
  const story = arm.engine.geometryReader.storyIds().find((story) => story.includes(':t') && arm.engine.paragraphs(story)[0]?.text === text);
  if (!story) throw new Error('Cell story missing');
  const paraId = arm.engine.paragraphs(story)[0]!.paraId;
  expect(execute(arm, command(arm, 'select', { selection: selection(paraId, 0, 0, story) })).status).toBe('noop');
  return arm;
}

function payload(arm: CommandArm): TablePayload | null {
  return arm.engine.tablePayload('body', 0) as unknown as TablePayload | null;
}

function layout(arm: CommandArm) {
  arm.engine.layoutDocumentWithRegionsRetained(JSON.stringify({
    bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } }, renderEnv: {},
  }));
  const frame = applyFrameDelta(null, decodeFrameDelta(arm.engine.buildDisplayListFrame('{}', 0)));
  return { frame, layoutRevision: arm.version };
}

function ordered(arm: CommandArm, intent: DocxCommandEnvelope, predecessor: DocxCommandSelection | null) {
  const readSelectionLayout = () => layout(arm);
  const ctx = context(arm, intent, predecessor);
  const checked = preflightDocxCommand(arm.engine, arm.state, intent, ctx, { readSelectionLayout });
  if (checked.status === 'refused') return checked;
  const result = applyDocxCommand(arm.engine, arm.state, checked.prepared, ctx);
  if (result.status === 'applied') arm.version++;
  return result;
}

function runAt(arm: CommandArm, text: string) {
  const initial = layout(arm);
  const page = initial.frame.pages[0]!;
  const run = page.page.primitives.find((primitive): primitive is TextRunPrimitive =>
    primitive.kind === 'text' && primitive.text === text);
  if (!run) throw new Error(`Rendered text is missing: ${text}`);
  return { initial, page, run };
}

function pointPayload(arm: CommandArm, text: string, shift = false) {
  const { initial, page, run } = runAt(arm, text);
  return {
    page: { pageId: page.pageId.toString(), pageIndex: 0 }, x: run.x + run.width / 2,
    y: run.baselineY - 5, layoutRevision: initial.layoutRevision, frameEpoch: initial.frame.frameEpoch,
    clickCount: 1, modifiers: { shift, alt: false, ctrl: false, meta: false },
  };
}

afterEach(closeCommandEngines);

beforeAll(loadRustDisplayListQueryEngine);

test('select_table_cells_covers_table_row_and_column_without_version_change', async () => {
  const arm = await inCell();
  const vector = arm.engine.encodeStateVector();
  const tabled = execute(arm, command(arm, 'selectTableCells', { target: 'table' }));
  expect(tabled.status).toBe('noop');
  expect(arm.version).toBe(0);
  expect(arm.engine.encodeStateVector()).toEqual(vector);
  expect(arm.engine.canUndo()).toBe(false);
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 1 } });
  expect(tabled.selection).toMatchObject({ anchor: { offset: 0 }, focus: { offset: 0 } });
  expect(tabled.editState?.tableContext).toMatchObject({
    rowIndex: 1, columnIndex: 1, rowCount: 2, columnCount: 2, hasMultiCellSelection: true, rowSpan: 1, columnSpan: 1,
  });
  const storyB = arm.engine.geometryReader.storyIds().find((story) => arm.engine.paragraphs(story)[0]?.text === 'B')!;
  execute(arm, command(arm, 'select', { selection: selection(arm.engine.paragraphs(storyB)[0]!.paraId, 0, 0, storyB) }));
  const row = execute(arm, command(arm, 'selectTableCells', { target: 'row' }));
  expect(row.status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 1 } });
  const storyC = arm.engine.geometryReader.storyIds().find((story) => arm.engine.paragraphs(story)[0]?.text === 'C')!;
  execute(arm, command(arm, 'select', { selection: selection(arm.engine.paragraphs(storyC)[0]!.paraId, 0, 0, storyC) }));
  expect(execute(arm, command(arm, 'selectTableCells', { target: 'column' })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 0 } });
}, TIMEOUT);

test('select_table_cells_reports_merged_spans_and_targets_the_innermost_table', async () => {
  const merged = paragraph('Before') + table([
    cell('Merged', '<w:gridSpan w:val="2"/>'),
    cell('C') + cell('D'),
  ]) + paragraph('After');
  const arm = await inCell(merged, 'Merged');
  const row = execute(arm, command(arm, 'selectTableCells', { target: 'row' }));
  expect(row.status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 0 } });
  expect(row.editState?.tableContext).toMatchObject({ rowIndex: 0, columnIndex: 0, columnSpan: 2, hasMultiCellSelection: false });
  expect(execute(arm, command(arm, 'selectTableCells', { target: 'table' })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 1 } });
  const nested = paragraph('Before') + table([
    cell('Outer') + `<w:tc>${paragraph('Host')}${table([cell('Inner A') + cell('Inner B'), cell('Inner C') + cell('Inner D')])}</w:tc>`,
  ]) + paragraph('After');
  const inner = await inCell(nested, 'Inner A');
  expect(inner.state.selection!.focus.story).not.toBe('body');
  expect(execute(inner, command(inner, 'selectTableCells', { target: 'table' })).status).toBe('noop');
  const range = inner.engine.cellSelection()!;
  expect(range.anchor.story).not.toBe('body');
  expect([range.anchor.row, range.anchor.column]).toEqual([0, 0]);
  expect([range.head.row, range.head.column]).toEqual([1, 1]);
}, TIMEOUT);

test('table_context_reports_spans_of_the_caret_cell_not_the_range_head', async () => {
  const merged = paragraph('Before') + table([
    cell('Merged', '<w:gridSpan w:val="2"/>'),
    cell('C') + cell('D'),
  ]) + paragraph('After');
  const arm = await inCell(merged, 'Merged');
  const selected = execute(arm, command(arm, 'selectTableCells', { target: 'table' }));
  expect(selected.status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 1 } });
  expect(selected.editState?.tableContext).toMatchObject({ rowIndex: 1, columnIndex: 1, columnSpan: 2, hasMultiCellSelection: true });
}, TIMEOUT);

test('select_table_cells_refuses_outside_tables_and_allows_read_only', async () => {
  const arm = await openEngine([], { bodyXml: body });
  const first = arm.engine.paragraphs('body')[0]!;
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'select', { selection: selection(first.paraId, 0) })).status).toBe('noop');
  const settled = snapshot(arm);
  const refused = execute(arm, command(arm, 'selectTableCells', { target: 'table' }));
  expect(refused).toMatchObject({ status: 'refused', reason: 'invalid-cell-target' });
  expect(snapshot(arm)).toEqual(settled);
  expect(before.version).toBe(0);
  const intent = command(arm, 'selectTableCells', { target: 'table' });
  expect(execute(arm, { ...intent, payload: { target: 'diagonal' } } as typeof intent))
    .toMatchObject({ status: 'refused', reason: 'invalid-payload' });
  const cellArm = await inCell();
  cellArm.state.readOnly = true;
  const vector = cellArm.engine.encodeStateVector();
  expect(execute(cellArm, command(cellArm, 'selectTableCells', { target: 'table' })).status).toBe('noop');
  expect(cellArm.engine.encodeStateVector()).toEqual(vector);
  expect(cellArm.engine.canUndo()).toBe(false);
  expect(cellArm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 1 } });
}, TIMEOUT);

test('pointer_drag_and_shift_click_span_cells_and_stay_collapsed_at_the_anchor', async () => {
  const arm = await inCell();
  const clicked = ordered(arm, command(arm, 'selectPoint', pointPayload(arm, 'A')), null);
  expect(clicked.status).not.toBe('refused');
  const anchorStory = arm.state.selection!.anchor.story;
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 0 } });
  const dragged = ordered(arm, command(arm, 'extendSelection', pointPayload(arm, 'D')), clicked.selection ?? null);
  expect(dragged.status).not.toBe('refused');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 1 } });
  expect(arm.state.selection!.anchor.story).toBe(anchorStory);
  expect(arm.state.selection!.anchor).toMatchObject(arm.state.selection!.focus);
  const shifted = ordered(arm, command(arm, 'selectPoint', pointPayload(arm, 'B', true)), dragged.selection ?? null);
  expect(shifted.status).not.toBe('refused');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 1 } });
}, TIMEOUT);

test('pointer_drag_into_a_nested_table_maps_to_the_outer_cell', async () => {
  const xml = paragraph('Before') + table([
    cell('A') + `<w:tc>${paragraph('Host')}${table([cell('Inner A') + cell('Inner B')])}</w:tc>`,
  ]) + paragraph('After');
  const arm = await inCell(xml);
  const clicked = ordered(arm, command(arm, 'selectPoint', pointPayload(arm, 'A')), null);
  expect(clicked.status).not.toBe('refused');
  const dragged = ordered(arm, command(arm, 'extendSelection', pointPayload(arm, 'Inner B')), clicked.selection ?? null);
  expect(dragged.status).not.toBe('refused');
  const range = arm.engine.cellSelection()!;
  expect(range.anchor.story).toBe('body');
  expect([range.anchor.row, range.anchor.column]).toEqual([0, 0]);
  expect([range.head.row, range.head.column]).toEqual([0, 1]);
}, TIMEOUT);

test('shift_arrow_steps_the_range_head_and_shrinks_back_to_one_cell', async () => {
  const arm = await inCell();
  const right = (extend: boolean) => command(arm, 'navigate', extend
    ? { unit: 'character', direction: 'forward', extend: true }
    : { unit: 'character', direction: 'forward', extend: false });
  expect(execute(arm, right(true)).status).toBe('noop');
  expect(arm.state.selection!.focus.offset).toBe(1);
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 0 } });
  expect(execute(arm, right(true)).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 1 } });
  expect(arm.state.selection!.anchor).toMatchObject(arm.state.selection!.focus);
  expect(execute(arm, command(arm, 'navigate', { unit: 'line', direction: 'down', extend: true, goalX: 0 })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 1 } });
  expect(execute(arm, command(arm, 'navigate', { unit: 'character', direction: 'backward', extend: true })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 0 } });
  expect(execute(arm, command(arm, 'navigate', { unit: 'line', direction: 'up', extend: true, goalX: 0 })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 0 } });
}, TIMEOUT);

test('plain_click_arrow_and_typing_clear_the_range_into_the_caret_cell', async () => {
  const arm = await inCell();
  execute(arm, command(arm, 'navigate', { unit: 'character', direction: 'forward', extend: true }));
  execute(arm, command(arm, 'navigate', { unit: 'character', direction: 'forward', extend: true }));
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 1 } });
  expect(execute(arm, command(arm, 'navigate', { unit: 'character', direction: 'forward', extend: false })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 0 } });
  execute(arm, command(arm, 'navigate', { unit: 'character', direction: 'forward', extend: true }));
  const typed = execute(arm, command(arm, 'insertText', { text: 'X' }));
  expect(typed.status).toBe('applied');
  const anchorStory = arm.state.selection!.anchor.story;
  expect(arm.engine.paragraphs(anchorStory)[0]!.text).toBe('AX');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 0 } });
  const first = arm.engine.paragraphs('body')[0]!;
  expect(execute(arm, command(arm, 'select', { selection: selection(first.paraId, 0) })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toBeNull();
}, TIMEOUT);

test('delete_backward_at_cell_start_clears_multi_cell_range_to_caret_cell', async () => {
  const arm = await inCell();
  expect(execute(arm, command(arm, 'selectTableCells', { target: 'row' })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 1 } });
  const deleted = execute(arm, command(arm, 'deleteBackward', {}));
  expect(deleted.status).toBe('noop');
  const range = arm.engine.cellSelection();
  expect(range).not.toBeNull();
  expect(range!.anchor).toMatchObject({ row: 0, column: 0 });
  expect(range!.head).toMatchObject({ row: 0, column: 0 });
  expect(range!.anchor.row).toBe(range!.head.row);
  expect(range!.anchor.column).toBe(range!.head.column);
  expect(deleted.editState?.tableContext?.hasMultiCellSelection).toBe(false);
}, TIMEOUT);

test('cell_selection_runtime_error_propagates_through_shift_arrow_range', async () => {
  const arm = await inCell();
  const base = arm.state.selection!;
  const moved = { ...base.focus, offset: base.focus.offset + 1 };
  Object.assign(arm.engine, {
    cellSelection: () => { throw new WebAssembly.RuntimeError('unreachable'); },
  });
  expect(() => resolveShiftArrowCellRange(arm.engine, base, moved, 'character', 'forward'))
    .toThrow(WebAssembly.RuntimeError);
  Object.assign(arm.engine, {
    cellSelection: () => { throw new TypeError('transient'); },
  });
  expect(resolveShiftArrowCellRange(arm.engine, base, moved, 'character', 'forward')).toBeNull();
}, TIMEOUT);

test('multi_cell_range_reports_one_rect_per_selected_cell', async () => {
  const arm = await inCell();
  expect(execute(arm, command(arm, 'selectTableCells', { target: 'table' })).status).toBe('noop');
  const { frame } = layout(arm);
  const rects = readResidentCellSelectionRects(arm.engine, frame, arm.engine.cellSelection());
  expect(rects?.length).toBe(4);
  for (const rect of rects!) {
    expect(rect.pageIndex).toBe(0);
    expect(rect.width).toBeGreaterThan(0);
    expect(rect.height).toBeGreaterThan(0);
  }
  expect(readResidentCellSelectionRects(arm.engine, frame, null)).toBeNull();
}, TIMEOUT);

test('select_row_from_a_middle_cell_then_merge_and_shade_apply_to_the_whole_row', async () => {
  const wide = paragraph('Before') + '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    + `<w:tr>${cell('A') + cell('B') + cell('C')}</w:tr><w:tr>${cell('D') + cell('E') + cell('F')}</w:tr></w:tbl>` + paragraph('After');
  const arm = await inCell(wide, 'B');
  expect(execute(arm, command(arm, 'selectTableCells', { target: 'row' })).status).toBe('noop');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 0, column: 2 } });
  const shaded = execute(arm, command(arm, 'setTableCellShading', { color: 'FF0000' }));
  expect(shaded.status).toBe('applied');
  expect(payload(arm)!.rows[0]!.cells.map((entry) => entry.tcPr?.backgroundColor)).toEqual(['FF0000', 'FF0000', 'FF0000']);
  expect(payload(arm)!.rows[1]!.cells.map((entry) => entry.tcPr?.backgroundColor)).toEqual([undefined, undefined, undefined]);
  expect(execute(arm, command(arm, 'selectTableCells', { target: 'row' })).status).toBe('noop');
  const merged = execute(arm, command(arm, 'mergeTableCells', {}));
  expect(merged.status).toBe('applied');
  expect(payload(arm)!.rows[0]!.cells).toHaveLength(1);
  expect(tableAnchors(payload(arm)!).anchors.filter((anchor) => anchor.row === 0)).toMatchObject([{ column: 0, colspan: 3 }]);
  expect(payload(arm)!.rows[1]!.cells).toHaveLength(3);
}, TIMEOUT);

test('merged_cell_range_reports_one_rect_per_distinct_cell_without_overlap', async () => {
  const merged = paragraph('Before') + table([
    cell('Merged', '<w:gridSpan w:val="2"/>'),
    cell('C') + cell('D'),
  ]) + paragraph('After');
  const arm = await inCell(merged, 'Merged');
  expect(execute(arm, command(arm, 'selectTableCells', { target: 'table' })).status).toBe('noop');
  const { frame } = layout(arm);
  const rects = readResidentCellSelectionRects(arm.engine, frame, arm.engine.cellSelection());
  expect(rects?.length).toBe(3);
  for (let first = 0; first < (rects?.length ?? 0); first += 1) {
    for (let second = first + 1; second < (rects?.length ?? 0); second += 1) {
      const left = rects![first]!;
      const right = rects![second]!;
      const overlaps = left.pageIndex === right.pageIndex &&
        left.x < right.x + right.width && right.x < left.x + left.width &&
        left.y < right.y + right.height && right.y < left.y + left.height;
      expect(overlaps).toBe(false);
    }
  }
}, TIMEOUT);

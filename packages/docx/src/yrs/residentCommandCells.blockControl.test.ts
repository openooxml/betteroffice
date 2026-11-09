import { afterEach, expect, test } from 'bun:test';
import { projectDocxEditState } from './residentCommandEditState';
import { closeCommandEngines, command, execute, openCommandEngine, selection, TIMEOUT } from './residentCommandTestSupport';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const control = (content: string) => `<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const cell = (content: string) => `<w:tc><w:tcPr/>${content}</w:tc>`;
const table = (rows: string[]) => `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
const body = paragraph('Before') + table([
  cell(paragraph('A')) + cell(control(paragraph('B'))),
  cell(control(paragraph('C')) + paragraph('C2')) + cell(paragraph('D')),
]) + paragraph('After');
const navigation = (direction: 'forward' | 'backward', extend = false) => ({ direction, extend, createLastRow: false });

type Arm = Awaited<ReturnType<typeof openCommandEngine>>;

function storyOf(arm: Arm, text: string): string {
  const story = arm.engine.geometryReader.storyIds().find((id) => arm.engine.paragraphs(id)[0]?.text === text);
  if (!story) throw new Error(`Story of ${text} missing`);
  return story;
}

async function at(text: string, editState = false) {
  const arm = await openCommandEngine([], { bodyXml: body, editState });
  const story = storyOf(arm, text);
  expect(execute(arm, command(arm, 'select', { selection: selection(arm.engine.paragraphs(story)[0]!.paraId, 0, 0, story) })).status).toBe('noop');
  return arm;
}

const focused = (arm: Arm) => arm.engine.paragraphs(arm.state.selection!.focus.story)[0]!.text;

afterEach(closeCommandEngines);

test('a_caret_inside_a_cell_block_control_has_table_context', async () => {
  const arm = await at('B', true);
  expect(arm.state.selection!.focus.story).toContain(':sdt');
  expect(projectDocxEditState(arm.engine, arm.state).tableContext).toMatchObject({ rowIndex: 0, columnIndex: 1 });
}, TIMEOUT);

test('tab_leaves_a_cell_block_control_in_both_directions', async () => {
  const forward = await at('B');
  expect(execute(forward, command(forward, 'navigateCell', navigation('forward'))).status).toBe('noop');
  expect(forward.state.selection!.focus.story).toContain(':sdt');
  expect(focused(forward)).toBe('C');
  const backward = await at('B');
  expect(execute(backward, command(backward, 'navigateCell', navigation('backward'))).status).toBe('noop');
  expect(focused(backward)).toBe('A');
}, TIMEOUT);

test('tab_enters_a_cell_at_its_first_block_through_a_block_control', async () => {
  const onlyControl = await at('A');
  expect(execute(onlyControl, command(onlyControl, 'navigateCell', navigation('forward'))).status).toBe('noop');
  expect(onlyControl.state.selection!.focus.story).toContain(':sdt');
  expect(focused(onlyControl)).toBe('B');
  expect(onlyControl.state.selection!.focus.offset).toBe(0);
  const controlThenParagraph = await at('D');
  expect(execute(controlThenParagraph, command(controlThenParagraph, 'navigateCell', navigation('backward'))).status).toBe('noop');
  expect(focused(controlThenParagraph)).toBe('C');
}, TIMEOUT);

test('navigating_out_of_a_block_control_changes_nothing_and_extends_cell_selection', async () => {
  const arm = await at('B');
  const before = arm.engine.encodeStateVector();
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward', true))).status).toBe('noop');
  expect(arm.engine.encodeStateVector()).toEqual(before);
  expect(arm.engine.cellSelection()).toMatchObject({
    anchor: { row: 0, column: 1 }, head: { row: 1, column: 0 },
  });
}, TIMEOUT);

test('a_row_inserted_from_inside_a_block_control_is_undoable_and_keeps_the_control', async () => {
  const arm = await at('B');
  expect(execute(arm, command(arm, 'insertTableRow', { side: 'below' })).status).toBe('applied');
  const rows = () => (arm.engine.tablePayload('body', 0) as { rows: unknown[] }).rows.length;
  expect(rows()).toBe(3);
  expect(arm.engine.paragraphs(storyOf(arm, 'B'))[0]!.text).toBe('B');
  expect(execute(arm, command(arm, 'undo', {})).status).toBe('applied');
  expect(rows()).toBe(2);
}, TIMEOUT);

const tableJson = (arm: Arm) => JSON.stringify(arm.engine.tablePayload('body', 0));

/** Applies a table command from inside the control, then expects undo to restore the table and the control's text. */
async function appliesAndUndoes(kind: string, payload: Record<string, unknown>, changed: (arm: Arm) => void, cells?: [number, number, number, number]) {
  const arm = await at('B');
  if (cells) arm.engine.setCellSelection({
    anchor: { story: 'body', tableIndex: 0, row: cells[0], column: cells[1] },
    head: { story: 'body', tableIndex: 0, row: cells[2], column: cells[3] },
  });
  const before = tableJson(arm);
  const version = arm.engine.proposalEngine.version();
  expect(execute(arm, command(arm, kind as 'insertTableRow', payload as never)).status).toBe('applied');
  expect(arm.engine.proposalEngine.version()).not.toEqual(version);
  changed(arm);
  expect(execute(arm, command(arm, 'undo', {})).status).toBe('applied');
  expect(tableJson(arm)).toBe(before);
  expect(arm.engine.paragraphs(storyOf(arm, 'B'))[0]!.text).toBe('B');
  expect(arm.engine.paragraphs(storyOf(arm, 'C'))[0]!.text).toBe('C');
}

const rows = (arm: Arm) => (arm.engine.tablePayload('body', 0) as unknown as { rows: { cells: { tcPr?: Record<string, unknown> }[] }[] }).rows;

test('deleting_a_row_from_inside_a_block_control_takes_that_row_and_undoes', async () => {
  await appliesAndUndoes('deleteTableRow', {}, (arm) => {
    expect(rows(arm)).toHaveLength(1);
    expect(arm.engine.geometryReader.storyIds().some((id) => arm.engine.paragraphs(id)[0]?.text === 'C')).toBe(true);
    expect(arm.engine.geometryReader.storyIds().some((id) => arm.engine.paragraphs(id)[0]?.text === 'A')).toBe(false);
  });
}, TIMEOUT);

test('merging_cells_from_inside_a_block_control_keeps_the_controls_text_and_undoes', async () => {
  await appliesAndUndoes('mergeTableCells', {}, (arm) => {
    expect(rows(arm)[0]!.cells).toHaveLength(1);
    const texts = arm.engine.geometryReader.storyIds().flatMap((id) => arm.engine.paragraphs(id).map((paragraph) => paragraph.text));
    expect(texts).toContain('A');
    expect(texts).toContain('B');
  }, [0, 0, 0, 1]);
}, TIMEOUT);

test('shading_from_inside_a_block_control_shades_its_cell_only_and_undoes', async () => {
  await appliesAndUndoes('setTableCellShading', { color: 'FF0000' }, (arm) => {
    const shaded = rows(arm).map((row) => row.cells.map((cell) => cell.tcPr?.backgroundColor === 'FF0000'));
    expect(shaded).toEqual([[false, true], [false, false]]);
  });
}, TIMEOUT);

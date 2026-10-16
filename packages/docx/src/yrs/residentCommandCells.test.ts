import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, command, execute, openCommandEngine, selection, snapshot, TIMEOUT } from './residentCommandTestSupport';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, properties = '') => `<w:tc><w:tcPr>${properties}</w:tcPr>${paragraph(text)}</w:tc>`;
const table = (rows: string[]) => `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
const body = paragraph('Before') + table([cell('A') + cell('B'), cell('C') + cell('D')]) + paragraph('After');
const navigation = (direction: 'forward' | 'backward', createLastRow = false, extend = false) => ({ direction, extend, createLastRow });

async function inCell(bodyXml = body, text = 'A') {
  const arm = await openCommandEngine([], { bodyXml });
  const story = arm.engine.geometryReader.storyIds().find((story) => story.includes(':t') && arm.engine.paragraphs(story)[0]?.text === text);
  if (!story) throw new Error('Cell story missing');
  const paraId = arm.engine.paragraphs(story)[0]!.paraId;
  expect(execute(arm, command(arm, 'select', { selection: selection(paraId, 0, 0, story) })).status).toBe('noop');
  return arm;
}

function focusedText(arm: Awaited<ReturnType<typeof inCell>>) {
  return arm.engine.paragraphs(arm.state.selection!.focus.story)[0]!.text;
}

afterEach(closeCommandEngines);

test('cell_navigation_wraps_rows_and_preserves_content_and_history', async () => {
  const arm = await inCell();
  const before = arm.engine.encodeStateVector();
  for (const text of ['B', 'C', 'D']) {
    expect(execute(arm, command(arm, 'navigateCell', navigation('forward'))).status).toBe('noop');
    expect(focusedText(arm)).toBe(text);
    expect(arm.state.selection!.focus.offset).toBe(0);
  }
  for (const text of ['C', 'B', 'A', 'A']) {
    expect(execute(arm, command(arm, 'navigateCell', navigation('backward'))).status).toBe('noop');
    expect(focusedText(arm)).toBe(text);
  }
  expect(arm.engine.encodeStateVector()).toEqual(before);
  expect(arm.engine.canUndo()).toBe(false);
  expect(arm.version).toBe(0);
}, TIMEOUT);

test('cell_navigation_skips_horizontal_and_vertical_merged_coverage', async () => {
  const xml = paragraph('Before') + table([
    cell('Merged', '<w:gridSpan w:val="2"/>'),
    cell('Vertical', '<w:vMerge w:val="restart"/>') + cell('B'),
    cell('', '<w:vMerge/>') + cell('C'),
  ]) + paragraph('After');
  const arm = await inCell(xml, 'Merged');
  for (const text of ['Vertical', 'B', 'C', 'C']) {
    expect(execute(arm, command(arm, 'navigateCell', navigation('forward'))).status).toBe('noop');
    expect(focusedText(arm)).toBe(text);
  }
  expect(execute(arm, command(arm, 'navigateCell', navigation('backward'))).status).toBe('noop');
  expect(focusedText(arm)).toBe('B');
}, TIMEOUT);

test('cell_navigation_rebases_from_predecessor_and_extends_cell_selection', async () => {
  const arm = await inCell();
  const firstCommand = command(arm, 'navigateCell', navigation('forward'));
  const first = execute(arm, firstCommand);
  const predecessor = first.selection!;
  expect(focusedText(arm)).toBe('B');
  const a = arm.engine.geometryReader.storyIds().find((story) => arm.engine.paragraphs(story)[0]?.text === 'A')!;
  execute(arm, command(arm, 'select', { selection: selection(arm.engine.paragraphs(a)[0]!.paraId, 0, 0, a) }));
  const next = execute(arm, command(arm, 'navigateCell', navigation('forward', false, true), { afterSeq: firstCommand.seq }), predecessor);
  expect(next.status).toBe('noop');
  expect(focusedText(arm)).toBe('C');
  expect(arm.engine.cellSelection()).toMatchObject({
    anchor: { row: 0, column: 1 }, head: { row: 1, column: 0 },
  });
}, TIMEOUT);

test('terminal_cell_row_creation_is_opt_in_readonly_guarded_and_undoable', async () => {
  const arm = await inCell(body, 'D');
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward'))).status).toBe('noop');
  expect(focusedText(arm)).toBe('D');
  arm.state.readOnly = true;
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward', true)))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(snapshot(arm)).toEqual({ ...before, readOnly: true });
  arm.state.readOnly = false;
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward', true))).status).toBe('applied');
  expect(focusedText(arm)).toBe('');
  const context = arm.engine.cellSelection()!.head;
  expect(context.row).toBe(2);
  expect(context.column).toBe(0);
  expect((arm.engine.tablePayload(context.story, context.tableIndex)!.rows as unknown[]).length).toBe(3);
  expect(arm.engine.canUndo()).toBe(true);
  arm.native.undo();
  expect((arm.engine.tablePayload(context.story, context.tableIndex)!.rows as unknown[]).length).toBe(2);
}, TIMEOUT);

test('cell_navigation_refuses_invalid_payload_before_any_mutation', async () => {
  const arm = await inCell();
  for (const payload of [
    { direction: 'sideways', extend: false, createLastRow: false },
    { direction: 'forward', extend: 'yes', createLastRow: false },
    { direction: 'forward', extend: false, createLastRow: 'yes' },
  ]) {
    const before = snapshot(arm);
    const intent = command(arm, 'navigateCell', navigation('forward'));
    expect(execute(arm, { ...intent, payload } as typeof intent)).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    expect(snapshot(arm)).toEqual(before);
  }
}, TIMEOUT);

test('readonly_cell_navigation_moves_without_creating_rows_or_history', async () => {
  const arm = await inCell();
  arm.state.readOnly = true;
  const vector = arm.engine.encodeStateVector();
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward', true))).status).toBe('noop');
  expect(focusedText(arm)).toBe('B');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 0, column: 1 }, head: { row: 0, column: 1 } });
  expect(arm.engine.encodeStateVector()).toEqual(vector);
  expect(arm.engine.canUndo()).toBe(false);
}, TIMEOUT);

test('nested_cell_navigation_stays_in_the_innermost_table', async () => {
  const xml = paragraph('Before') + table([
    cell('Outer') + `<w:tc>${paragraph('Host')}${table([cell('Inner A') + cell('Inner B')])}</w:tc>`,
  ]) + paragraph('After');
  const arm = await inCell(xml, 'Inner A');
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward'))).status).toBe('noop');
  expect(focusedText(arm)).toBe('Inner B');
  expect(arm.engine.cellSelection()!.head.story).toContain(':r0c1');
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward'))).status).toBe('noop');
  expect(focusedText(arm)).toBe('Inner B');
}, TIMEOUT);

test('terminal_merged_row_creates_a_surviving_cell_target', async () => {
  const xml = paragraph('Before') + table([
    cell('A', '<w:vMerge w:val="restart"/>') + cell('B'),
    cell('', '<w:vMerge/>') + cell('C'),
  ]) + paragraph('After');
  const arm = await inCell(xml, 'C');
  expect(execute(arm, command(arm, 'navigateCell', navigation('forward', true, true))).status).toBe('applied');
  expect(focusedText(arm)).toBe('');
  expect(arm.engine.cellSelection()).toMatchObject({ anchor: { row: 1, column: 1 }, head: { row: 2, column: 0 } });
}, TIMEOUT);

test('cell_navigation_after_leaving_a_table_preserves_following_input', async () => {
  const arm = await inCell();
  const paragraph = arm.engine.paragraphs('body')[0]!;
  const selected = command(arm, 'select', { selection: selection(paragraph.paraId, 0) });
  const outside = execute(arm, selected);
  const next = command(arm, 'navigateCell', navigation('forward', true), { afterSeq: selected.seq });
  const navigated = execute(arm, next, outside.selection);
  expect(navigated.status).toBe('noop');
  expect(navigated.selection?.focus.story).toBe('body');
  expect(execute(arm, command(arm, 'insertText', { text: 'valid' }, { afterSeq: next.seq }), navigated.selection).status).toBe('applied');
  expect(arm.engine.paragraphs('body')[0]!.text).toBe('validBefore');
}, TIMEOUT);

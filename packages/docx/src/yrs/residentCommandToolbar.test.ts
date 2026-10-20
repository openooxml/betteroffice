import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, command, execute, openCommandEngine, selection, snapshot, texts, TIMEOUT } from './residentCommandTestSupport';

afterEach(closeCommandEngines);

test('worker_selection_formatting_and_list_commands_cover_multiple_paragraphs', async () => {
  const arm = await openCommandEngine(['Alpha', 'Beta']);
  const [first, second] = arm.engine.paragraphs('body');
  const selected = { anchor: { story: 'body', paraId: first!.paraId, offset: 0 }, focus: { story: 'body', paraId: second!.paraId, offset: 4 } };
  execute(arm, command(arm, 'select', { selection: selected }));
  expect(execute(arm, command(arm, 'formatSelection', { delta: { highlight: 'yellow', strike: true, fontSize: 14 } }))).toMatchObject({ status: 'applied', editState: { toolbar: { strike: true, highlight: 'yellow' } } });
  expect(execute(arm, command(arm, 'toggleList', { kind: 'bullet' }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.paragraphs('body').map(p => p.properties.numPr)).toEqual([{ numId: 1, ilvl: 0 }, { numId: 1, ilvl: 0 }]);
  expect(execute(arm, command(arm, 'changeIndent', { direction: 'increase' }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.paragraphs('body').map(p => p.properties.numPr)).toEqual([{ numId: 1, ilvl: 1 }, { numId: 1, ilvl: 1 }]);
  execute(arm, command(arm, 'toggleList', { kind: 'bullet' }));
  expect(arm.engine.paragraphs('body').every(p => p.properties.numPr == null)).toBe(true);
  expect(execute(arm, command(arm, 'changeIndent', { direction: 'increase' }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.paragraphs('body').map(p => p.properties.indentLeft)).toEqual([720, 720]);
  expect(execute(arm, command(arm, 'changeIndent', { direction: 'decrease' }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.paragraphs('body').map(p => p.properties.indentLeft)).toEqual([0, 0]);
  execute(arm, command(arm, 'changeIndent', { direction: 'decrease' }));
  expect(arm.engine.paragraphs('body').map(p => p.properties.indentLeft)).toEqual([0, 0]);
}, TIMEOUT);

test('collapsed_toolbar_marks_follow_insertions_and_insertion_refusals_are_atomic', async () => {
  const arm = await openCommandEngine(['Alpha']);
  expect(execute(arm, command(arm, 'toggleStrike', {}))).toMatchObject({ status: 'noop', editState: { toolbar: { strike: true } } });
  execute(arm, command(arm, 'insertText', { text: 'X' }));
  expect(arm.engine.storySegments('body')[0]?.attributes.strike).toEqual({ double: false });
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'insertTable', { rows: 0, columns: 2 }))).toMatchObject({ status: 'refused', reason: 'invalid-table-size' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, command(arm, 'insertSectionBreak', { type: 'continuous', suggesting: { name: 'Ada', date: '2026-01-01' } }))).toMatchObject({ status: 'refused', reason: 'unsupported-suggesting' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, command(arm, 'insertImage', { image: { rId: 'missing', widthEmu: -1 } }))).toMatchObject({ status: 'refused', reason: 'invalid-image' });
  expect(snapshot(arm)).toEqual(before);
}, TIMEOUT);

test('worker_insertions_replace_selection_and_create_worker_owned_cells', async () => {
  const arm = await openCommandEngine(['Alpha']);
  const paraId = arm.state.selection!.focus.paraId;
  expect(execute(arm, command(arm, 'insertImage', { selection: selection(paraId, 1, 4), image: { rId: 'image', widthEmu: 952500, heightEmu: 952500 } }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.storySegments('body').some(s => s.kind === 'embed' && s.embedKind === 'image')).toBe(true);
  expect(execute(arm, command(arm, 'undo', {}))).toMatchObject({ status: 'applied' });
  expect(texts(arm)).toEqual(['Alpha']);
  expect(execute(arm, command(arm, 'insertTable', { rows: 2, columns: 2, selection: selection(paraId, 5) }))).toMatchObject({ status: 'applied' });
  expect(arm.state.selection?.focus.story).toMatch(/:t\d+:r0c0~\d+$/);
  expect(arm.engine.paragraphs(arm.state.selection!.focus.story)).toHaveLength(1);
}, TIMEOUT);

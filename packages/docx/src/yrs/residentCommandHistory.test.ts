import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, command, execute, openCommandEngine, selection, snapshot, texts, TIMEOUT } from './residentCommandTestSupport';

afterEach(closeCommandEngines);

test('history_commands_noop_without_history_and_refuse_payloads_before_mutation', async () => {
  const arm = await openCommandEngine();
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'undo', {}))).toMatchObject({ status: 'noop', editState: { canUndo: false, canRedo: false } });
  expect(execute(arm, command(arm, 'redo', {}))).toMatchObject({ status: 'noop' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, { ...command(arm, 'undo', {}), payload: { selection: arm.state.selection! } })).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
  expect(snapshot(arm)).toEqual(before);
}, TIMEOUT);

test('history_commands_preserve_paragraph_identity_and_repair_selection', async () => {
  const arm = await openCommandEngine(['Alpha']);
  const id = arm.state.selection!.focus.paraId;
  execute(arm, command(arm, 'paragraphBreak', { selection: selection(id, 5) }));
  const second = arm.engine.paragraphs('body')[1]!.paraId;
  expect(execute(arm, command(arm, 'undo', {}))).toMatchObject({ status: 'applied' });
  expect(texts(arm)).toEqual(['Alpha']);
  expect(arm.state.selection?.focus.paraId).toBe(id);
  expect(execute(arm, command(arm, 'redo', {}))).toMatchObject({ status: 'applied' });
  expect(arm.engine.paragraphs('body').map(p => p.paraId)).toEqual([id, second]);
}, TIMEOUT);

import { afterEach, expect, spyOn, test } from 'bun:test';
import { closeCommandEngines, command, execute, openCommandEngine, selection, snapshot, texts, TIMEOUT, type CommandArm } from './residentCommandTestSupport';

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

type InputTimes = { inputTime?: number; lastInputTime?: number };

function typeWithApplyDelay(arm: CommandArm, keys: [string, InputTimes][], applyMs: number) {
  let now = 1_000_000;
  const clock = spyOn(Date, 'now').mockImplementation(() => now);
  try {
    for (const [text, times] of keys) {
      execute(arm, command(arm, 'insertText', { text }, times));
      now += applyMs;
    }
  } finally {
    clock.mockRestore();
  }
}

function undoTexts(arm: CommandArm, steps: number) {
  return Array.from({ length: steps }, () => {
    execute(arm, command(arm, 'undo', {}));
    return texts(arm)[0];
  });
}

test('typing_applied_late_is_one_undo_step_when_typed_within_the_capture_window', async () => {
  const arm = await openCommandEngine(['Alpha']);
  typeWithApplyDelay(arm, [['B', { inputTime: 10 }], ['e', { inputTime: 60 }], ['t', { inputTime: 110 }], ['a', { inputTime: 160 }]], 700);
  expect(texts(arm)).toEqual(['BetaAlpha']);
  expect(undoTexts(arm, 1)).toEqual(['Alpha']);
}, TIMEOUT);

test('typing_pauses_split_undo_steps_when_edits_apply_together', async () => {
  const arm = await openCommandEngine(['Alpha']);
  typeWithApplyDelay(arm, [['B', { inputTime: 10 }], ['e', { inputTime: 60 }], ['t', { inputTime: 700 }], ['a', { inputTime: 750 }]], 0);
  expect(undoTexts(arm, 2)).toEqual(['BeAlpha', 'Alpha']);
}, TIMEOUT);

test('coalesced_typing_groups_from_its_last_input', async () => {
  const arm = await openCommandEngine(['Alpha']);
  typeWithApplyDelay(arm, [['Be', { inputTime: 10, lastInputTime: 610 }], ['ta', { inputTime: 700, lastInputTime: 760 }], ['!', { inputTime: 1_400 }]], 700);
  expect(undoTexts(arm, 2)).toEqual(['BetaAlpha', 'Alpha']);
}, TIMEOUT);

test('edits_without_input_time_keep_apply_time_grouping', async () => {
  const arm = await openCommandEngine(['Alpha']);
  typeWithApplyDelay(arm, [['B', {}], ['e', {}]], 700);
  expect(undoTexts(arm, 1)).toEqual(['BAlpha']);
}, TIMEOUT);

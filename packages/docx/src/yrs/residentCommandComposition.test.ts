import { afterEach, expect, spyOn, test } from 'bun:test';
import {
  closeCommandEngines, command, execute, expectCaret, openCommandEngine, selection, snapshot, storyRange, texts, TIMEOUT,
  type CommandArm,
} from './residentCommandTestSupport';

afterEach(closeCommandEngines);

const WRITER = { name: 'Writer', date: '2026-11-17T00:00:00Z' };

function step(arm: CommandArm, text: string, replaceBefore?: string, extra: Record<string, unknown> = {}) {
  return execute(arm, command(arm, 'insertText', { text, ...(replaceBefore === undefined ? {} : { replaceBefore }), ...extra }));
}

test('composition_steps_replace_the_text_before_the_caret_and_keep_stored_marks', async () => {
  const arm = await openCommandEngine(['abc']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'toggleMark', { mark: { type: 'bold' } })).status).toBe('noop');
  expect(step(arm, 'ㄱ').status).toBe('applied');
  const replaced = step(arm, '가', 'ㄱ');
  expect(replaced.status).toBe('applied');
  expect(texts(arm)).toEqual(['가abc']);
  expectCaret(replaced, { story: 'body', paraId: id, offset: 1 });
  expect(step(arm, '간', '가').status).toBe('applied');
  expect(texts(arm)).toEqual(['간abc']);
  expect(arm.engine.selectionContext(storyRange(selection(id, 0, 1))).bold).toBe(true);
  expect(arm.engine.selectionContext(storyRange(selection(id, 1, 2))).bold).not.toBe(true);
  expect(arm.state.storedMarks?.delta.bold).toBe(true);
  expect(step(arm, 'X').status).toBe('applied');
  expect(arm.engine.selectionContext(storyRange(selection(id, 1, 2))).bold).toBe(true);
}, TIMEOUT);

test('composition_step_is_refused_without_writing_when_other_text_is_before_the_caret', async () => {
  const arm = await openCommandEngine(['abc']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  execute(arm, command(arm, 'select', { selection: selection(id, 2) }));
  const before = snapshot(arm);
  expect(step(arm, '가', 'ㄱ')).toMatchObject({ status: 'refused', reason: 'replaced-text-mismatch' });
  expect(step(arm, '가', 'xab')).toMatchObject({ status: 'refused', reason: 'replaced-text-mismatch' });
  expect(execute(arm, command(arm, 'insertText', { text: '가', replaceBefore: 'ab', selection: selection(id, 1, 2) })))
    .toMatchObject({ status: 'refused', reason: 'replaced-text-mismatch' });
  expect(snapshot(arm)).toEqual(before);
  expect(step(arm, '가', 'ab').status).toBe('applied');
  expect(texts(arm)).toEqual(['가c']);
}, TIMEOUT);

test('composition_step_payloads_are_validated', async () => {
  const arm = await openCommandEngine(['abc']);
  const before = snapshot(arm);
  expect(step(arm, '가', '')).toMatchObject({ status: 'refused', reason: 'invalid-text' });
  expect(step(arm, '가', 'a\nb')).toMatchObject({ status: 'refused', reason: 'invalid-text' });
  expect(step(arm, '가', '￼')).toMatchObject({ status: 'refused', reason: 'invalid-text' });
  expect(execute(arm, { ...command(arm, 'insertText', { text: '가' }), payload: { text: '가', replaceBefore: 1 } as never }))
    .toMatchObject({ status: 'refused', reason: 'invalid-text' });
  expect(execute(arm, { ...command(arm, 'replaceSelection', { text: '가' }), payload: { text: '가', replaceBefore: 'a' } as never }))
    .toMatchObject({ status: 'refused', reason: 'invalid-payload' });
  expect(snapshot(arm)).toEqual(before);
}, TIMEOUT);

test('an_empty_composition_step_removes_the_composed_text', async () => {
  const arm = await openCommandEngine(['abc']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  step(arm, 'ㄱ');
  step(arm, '가', 'ㄱ');
  const removed = step(arm, '', '가');
  expect(removed.status).toBe('applied');
  expect(texts(arm)).toEqual(['abc']);
  expectCaret(removed, { story: 'body', paraId: id, offset: 0 });
}, TIMEOUT);

test('composition_steps_replace_a_surrogate_pair_whole', async () => {
  const arm = await openCommandEngine(['abc']);
  step(arm, '\u{20000}');
  expect(step(arm, '가', '\u{20000}').status).toBe('applied');
  expect(texts(arm)).toEqual(['가abc']);
}, TIMEOUT);

test('one_composition_is_one_undo_step_however_long_it_is_composed', async () => {
  const arm = await openCommandEngine(['Alpha']);
  let now = 1_000_000;
  const clock = spyOn(Date, 'now').mockImplementation(() => now);
  try {
    const steps: Array<[string, string | undefined, number, number]> = [
      ['ㅎ', undefined, 10, 10], ['하', 'ㅎ', 10, 700], ['한', '하', 10, 1_400],
      ['ㄱ', undefined, 2_100, 2_100], ['그', 'ㄱ', 2_100, 2_800], ['글', '그', 2_100, 3_500],
    ];
    for (const [text, replaceBefore, inputTime, lastInputTime] of steps) {
      execute(arm, command(arm, 'insertText', { text, ...(replaceBefore ? { replaceBefore } : {}) }, { inputTime, lastInputTime }));
      now += 700;
    }
  } finally {
    clock.mockRestore();
  }
  expect(texts(arm)).toEqual(['한글Alpha']);
  execute(arm, command(arm, 'undo', {}));
  expect(texts(arm)).toEqual(['한Alpha']);
  execute(arm, command(arm, 'undo', {}));
  expect(texts(arm)).toEqual(['Alpha']);
}, TIMEOUT);

test('suggested_composition_steps_remove_their_own_insertion_instead_of_striking_it', async () => {
  const arm = await openCommandEngine(['abc']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'insertText', { text: 'ㄱ', selection: selection(id, 1, 2), suggesting: WRITER })).status).toBe('applied');
  expect(step(arm, '가', 'ㄱ', { suggesting: WRITER }).status).toBe('applied');
  expect(step(arm, '각', '가', { suggesting: WRITER }).status).toBe('applied');
  expect(texts(arm)).toEqual(['ab각c']);
  const removed = step(arm, '', '각', { suggesting: WRITER });
  expect(removed.status).toBe('applied');
  expect(texts(arm)).toEqual(['abc']);
}, TIMEOUT);

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

test('composition_steps_take_the_resident_text_edit_and_keep_the_typed_over_formatting', async () => {
  const arm = await openCommandEngine(['abc def']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'toggleMark', { mark: { type: 'bold' }, selection: selection(id, 4, 7) })).status).toBe('applied');
  const resident = spyOn(arm.native, 'edit_text_resident');
  const replaced = spyOn(arm.native, 'replace_range');
  const spans = spyOn(arm.engine.geometryReader, 'paragraphSpans');
  try {
    // Typing over a bold word: the composition keeps the bold although the space before it is not bold.
    expect(execute(arm, command(arm, 'replaceSelection', { text: 'ㄱ', selection: selection(id, 4, 7) })).status).toBe('applied');
    expect(step(arm, '가', 'ㄱ').status).toBe('applied');
    expect(step(arm, '각', '가').status).toBe('applied');
    expect(texts(arm)).toEqual(['abc 각']);
    expect(arm.engine.selectionContext(storyRange(selection(id, 4, 5))).bold).toBe(true);
    expect(replaced).not.toHaveBeenCalled();
    expect(resident.mock.calls.map(call => [call[2], call[4], call[5]])).toEqual([[4, 7, 'ㄱ'], [4, 5, '가'], [4, 5, '각']]);
    // A step reads the caret's paragraph alone, not every paragraph of the story.
    expect(spans).not.toHaveBeenCalled();
    resident.mockClear();
    const removed = step(arm, '', '각');
    expect(removed.status).toBe('applied');
    expect(texts(arm)).toEqual(['abc ']);
    expectCaret(removed, { story: 'body', paraId: id, offset: 4 });
    expect(resident.mock.calls.map(call => [call[2], call[4], call[5]])).toEqual([[4, 5, undefined]]);
    expect(replaced).not.toHaveBeenCalled();
  } finally {
    spans.mockRestore();
    resident.mockRestore();
    replaced.mockRestore();
  }
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

/** This editor and a collaborator on the same document; `sync` sends one side's missing edits to the other. */
async function collaborators(paragraphs: string[]) {
  const arm = await openCommandEngine(paragraphs, { clientId: 8301 });
  const peer = await openCommandEngine(paragraphs, { clientId: 8302, checkpoint: arm.engine.encodeState() });
  const sync = (to: CommandArm, from: CommandArm) => to.engine.applyUpdate(from.engine.encodeStateAsUpdate(to.engine.encodeStateVector()));
  return { arm, peer, id: arm.engine.paragraphs('body')[0]!.paraId, sync };
}

test('a_composition_types_over_its_own_text_when_a_collaborator_types_right_after_it', async () => {
  const { arm, peer, id, sync } = await collaborators(['abc']);
  execute(arm, command(arm, 'select', { selection: selection(id, 1) }));
  expect(step(arm, 'ㄱ').status).toBe('applied');
  sync(peer, arm);
  // Exactly at the caret, between the composed text and what followed it.
  peer.engine.insertText({ story: 'body', paraId: id, offset: 2 }, 'X');
  sync(arm, peer);
  expect(texts(arm)).toEqual(['aㄱXbc']);
  const typed = step(arm, '가', 'ㄱ');
  expect(typed.status).toBe('applied');
  expect(texts(arm)).toEqual(['a가Xbc']);
  expectCaret(typed, { story: 'body', paraId: id, offset: 2 });
  expect(step(arm, '각', '가').status).toBe('applied');
  sync(peer, arm);
  expect(texts(arm)).toEqual(['a각Xbc']);
  expect(texts(peer)).toEqual(['a각Xbc']);
}, TIMEOUT);

test('a_composition_keeps_text_a_collaborator_types_before_it_and_further_on', async () => {
  const { arm, peer, id, sync } = await collaborators(['abc']);
  execute(arm, command(arm, 'select', { selection: selection(id, 1) }));
  expect(step(arm, 'ㅎ').status).toBe('applied');
  sync(peer, arm);
  peer.engine.insertText({ story: 'body', paraId: id, offset: 1 }, 'P');
  peer.engine.insertText({ story: 'body', paraId: id, offset: 4 }, 'Q');
  sync(arm, peer);
  expect(texts(arm)).toEqual(['aPㅎbQc']);
  expect(step(arm, '하', 'ㅎ').status).toBe('applied');
  expect(step(arm, '한', '하').status).toBe('applied');
  expect(texts(arm)).toEqual(['aP한bQc']);
}, TIMEOUT);

test('a_composition_step_is_refused_and_writes_nothing_when_a_collaborator_deleted_the_composed_text', async () => {
  // The text before the deleted composition reads the same as it: comparing text alone would type over it.
  const { arm, peer, id, sync } = await collaborators(['하bc']);
  execute(arm, command(arm, 'select', { selection: selection(id, 1) }));
  expect(step(arm, '하').status).toBe('applied');
  sync(peer, arm);
  peer.engine.deleteRange({ story: 'body', start: { paraId: id, offset: 1 }, end: { paraId: id, offset: 2 } });
  sync(arm, peer);
  expect(texts(arm)).toEqual(['하bc']);
  const before = snapshot(arm);
  expect(step(arm, '한', '하')).toMatchObject({ status: 'refused', reason: 'replaced-text-mismatch' });
  expect(snapshot(arm)).toEqual(before);
}, TIMEOUT);

test('a_composition_step_is_refused_when_a_collaborator_typed_inside_the_composed_text', async () => {
  const { arm, peer, id, sync } = await collaborators(['abc']);
  execute(arm, command(arm, 'select', { selection: selection(id, 3) }));
  expect(step(arm, 'か').status).toBe('applied');
  expect(step(arm, 'かん', 'か').status).toBe('applied');
  sync(peer, arm);
  peer.engine.insertText({ story: 'body', paraId: id, offset: 4 }, 'X');
  sync(arm, peer);
  expect(texts(arm)).toEqual(['abcかXん']);
  const before = snapshot(arm);
  expect(step(arm, 'かんじ', 'かん')).toMatchObject({ status: 'refused', reason: 'replaced-text-mismatch' });
  expect(snapshot(arm)).toEqual(before);
}, TIMEOUT);

test('a_composition_types_over_its_text_after_a_collaborator_formats_it', async () => {
  const { arm, peer, id, sync } = await collaborators(['abc']);
  execute(arm, command(arm, 'select', { selection: selection(id, 3) }));
  expect(step(arm, 'ㄱ').status).toBe('applied');
  sync(peer, arm);
  peer.engine.formatRange({ story: 'body', start: { paraId: id, offset: 3 }, end: { paraId: id, offset: 4 } }, { bold: true });
  sync(arm, peer);
  expect(step(arm, '가', 'ㄱ').status).toBe('applied');
  expect(texts(arm)).toEqual(['abc가']);
}, TIMEOUT);

test('a_composition_is_one_undo_step_with_a_collaborator_edit_between_its_steps', async () => {
  const { arm, peer, id, sync } = await collaborators(['Alpha']);
  let now = 1_000_000;
  const clock = spyOn(Date, 'now').mockImplementation(() => now);
  try {
    execute(arm, command(arm, 'insertText', { text: 'ㅎ' }, { inputTime: 10, lastInputTime: 10 }));
    now += 700;
    sync(peer, arm);
    peer.engine.insertText({ story: 'body', paraId: id, offset: 1 }, 'X');
    sync(arm, peer);
    now += 700;
    execute(arm, command(arm, 'insertText', { text: '하', replaceBefore: 'ㅎ' }, { inputTime: 10, lastInputTime: 1_400 }));
    now += 700;
    execute(arm, command(arm, 'insertText', { text: '한', replaceBefore: '하' }, { inputTime: 10, lastInputTime: 2_100 }));
  } finally {
    clock.mockRestore();
  }
  expect(texts(arm)).toEqual(['한XAlpha']);
  execute(arm, command(arm, 'undo', {}));
  expect(texts(arm)).toEqual(['XAlpha']);
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

test('keys_the_engine_cannot_type_with_stored_marks_are_typed_then_formatted', async () => {
  // An engine that has not laid the document out takes no resident edit.
  const arm = await openCommandEngine(['abc']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'toggleMark', { mark: { type: 'bold' } })).status).toBe('noop');
  const marked = spyOn(arm.native, 'edit_text_resident_marked').mockReturnValue(undefined);
  const formatted = spyOn(arm.native, 'format_range');
  try {
    expect(step(arm, 'x').status).toBe('applied');
    expect(texts(arm)).toEqual(['xabc']);
    expect(arm.engine.selectionContext(storyRange(selection(id, 0, 1))).bold).toBe(true);
    expect(marked).toHaveBeenCalledTimes(1);
    expect(formatted).toHaveBeenCalledTimes(1);
  } finally {
    formatted.mockRestore();
    marked.mockRestore();
  }
  // Suggested typing keeps the two edits: the formatting is tracked as its own change.
  const suggestedMarked = spyOn(arm.native, 'edit_text_resident_marked');
  const suggestedFormatted = spyOn(arm.native, 'format_range');
  try {
    expect(step(arm, 'y', undefined, { suggesting: WRITER }).status).toBe('applied');
    expect(suggestedMarked).not.toHaveBeenCalled();
    expect(suggestedFormatted).toHaveBeenCalledTimes(1);
  } finally {
    suggestedFormatted.mockRestore();
    suggestedMarked.mockRestore();
  }
}, TIMEOUT);

function timed(arm: CommandArm, kind: 'insertText' | 'replaceSelection', payload: Record<string, unknown>, inputTime: number, lastInputTime = inputTime) {
  return execute(arm, command(arm, kind, payload as never, { inputTime, lastInputTime }));
}

test('cancelling_a_composition_over_a_selection_gives_the_selection_back_with_its_formatting', async () => {
  const arm = await openCommandEngine(['abc def']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'toggleMark', { mark: { type: 'bold' }, selection: selection(id, 4, 7) })).status).toBe('applied');
  expect(timed(arm, 'insertText', { text: 'Z', selection: selection(id, 0) }, 10).status).toBe('applied');
  expect(timed(arm, 'replaceSelection', { text: 'ㄱ', selection: selection(id, 5, 8) }, 100).status).toBe('applied');
  expect(timed(arm, 'insertText', { text: '가', replaceBefore: 'ㄱ' }, 100, 200).status).toBe('applied');
  expect(texts(arm)).toEqual(['Zabc 가']);
  expect(timed(arm, 'insertText', { text: '', replaceBefore: '가' }, 100, 300).status).toBe('applied');
  expect(texts(arm)).toEqual(['Zabc def']);
  expect(arm.engine.selectionContext(storyRange(selection(id, 5, 8))).bold).toBe(true);
  // The typing before the composition is a step of its own.
  execute(arm, command(arm, 'undo', {}));
  expect(texts(arm)).toEqual(['abc def']);
}, TIMEOUT);

test('a_composition_over_a_selection_is_one_undo_step_apart_from_the_typing_before_it', async () => {
  const arm = await openCommandEngine(['abc def']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  timed(arm, 'insertText', { text: 'Z', selection: selection(id, 0) }, 10);
  timed(arm, 'replaceSelection', { text: 'ㄱ', selection: selection(id, 5, 8) }, 100);
  timed(arm, 'insertText', { text: '가', replaceBefore: 'ㄱ' }, 100, 200);
  expect(texts(arm)).toEqual(['Zabc 가']);
  execute(arm, command(arm, 'undo', {}));
  expect(texts(arm)).toEqual(['Zabc def']);
  execute(arm, command(arm, 'undo', {}));
  expect(texts(arm)).toEqual(['abc def']);
}, TIMEOUT);

test('cancelling_a_composition_over_a_selection_keeps_what_a_collaborator_typed_inside_it_and_elsewhere', async () => {
  const { arm, peer, id, sync } = await collaborators(['abc def']);
  timed(arm, 'replaceSelection', { text: 'x', selection: selection(id, 4, 7) }, 100);
  timed(arm, 'insertText', { text: 'xy', replaceBefore: 'x' }, 100, 200);
  sync(peer, arm);
  expect(execute(peer, command(peer, 'insertText', { text: 'Q', selection: selection(id, 5) })).status).toBe('applied');
  expect(execute(peer, command(peer, 'insertText', { text: 'P', selection: selection(id, 0) })).status).toBe('applied');
  sync(arm, peer);
  expect(texts(arm)).toEqual(['Pabc xQy']);
  expect(timed(arm, 'insertText', { text: '', replaceBefore: 'xy' }, 100, 300).status).toBe('applied');
  sync(peer, arm);
  expect(texts(arm)[0]).toMatch(/^Pabc (Qdef|defQ)$/);
  expect(texts(peer)).toEqual(texts(arm));
}, TIMEOUT);

test('an_empty_step_without_an_input_time_removes_text_typed_over_a_selection_instead_of_undoing', async () => {
  const arm = await openCommandEngine(['abc def']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  execute(arm, command(arm, 'replaceSelection', { text: 'ㄱ', selection: selection(id, 4, 7) }));
  expect(step(arm, '', 'ㄱ').status).toBe('applied');
  expect(texts(arm)).toEqual(['abc ']);
}, TIMEOUT);

import { afterEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createCommandHost } from '../../../../shared/office-session/commands/host';
import { isCommandOutcomeUnknownError } from '../../../../shared/office-session/commands/types';
import { isTransferResult } from '../../../../shared/office-session/protocol';
import { resolveImageLayoutAttrs, type ImageLayoutCurrentAttrs, type ImageLayoutTarget } from '../docx/imageLayout';
import { repackDocx } from '../docx/rezip';
import { unzipContainer } from '../docx/wasm';
import { preloadParseWasm } from '../wasm/parse';
import { applyDocxCommand, createResidentDocxCommandHooks, preflightDocxCommand } from './residentCommandActions';
import { createResidentCommandHost } from './residentCommandHost';
import { resolveCommandSelection } from './residentCommandAnchors';
import { readResidentProperties } from './residentProperties';
import { wrapSession } from './yrsSessionFacade';
import { yrsToDocument } from './yrsToDocument';
import type {
  DocxCommandEnvelope, DocxCommandIntent, DocxCommandPayload, DocxCommandSelection, DocxEditState,
} from './residentCommandProtocol';
import {
  TIMEOUT, closeCommandEngines, command, context, execute, expectCaret,
  openCommandEngine, selection, snapshot, storyRange, texts,
} from './residentCommandTestSupport';

afterEach(closeCommandEngines);

test('d1a_types_replaces_deletes_and_breaks_without_peer', async () => {
  const arm = await openCommandEngine();
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const secondId = arm.engine.paragraphs('body')[1]!.paraId;
  const inserted = execute(arm, command(arm, 'insertText', { text: '!', selection: selection(id, 10) }));
  expect(inserted.status).toBe('applied');
  expect(texts(arm)).toEqual(['Alpha beta!', 'Second']);
  expectCaret(inserted, { story: 'body', paraId: id, offset: 11 });
  expect(arm.version).toBe(1);

  const replaced = execute(arm, command(arm, 'replaceSelection', { text: 'WORLD', selection: selection(id, 11, 6) }));
  expect(replaced.status).toBe('applied');
  expect(texts(arm)).toEqual(['Alpha WORLD', 'Second']);
  expectCaret(replaced, { story: 'body', paraId: id, offset: 11 });
  const backward = execute(arm, command(arm, 'deleteBackward', {}));
  expect(backward.status).toBe('applied');
  expect(texts(arm)).toEqual(['Alpha WORL', 'Second']);
  expectCaret(backward, { story: 'body', paraId: id, offset: 10 });
  const forward = execute(arm, command(arm, 'deleteForward', { selection: selection(id, 6) }));
  expect(forward.status).toBe('applied');
  expect(texts(arm)).toEqual(['Alpha ORL', 'Second']);
  expectCaret(forward, { story: 'body', paraId: id, offset: 6 });

  const broken = execute(arm, command(arm, 'paragraphBreak', { selection: selection(id, 6, 9) }));
  expect(broken.status).toBe('applied');
  const paragraphs = arm.engine.paragraphs('body');
  expect(paragraphs.map(({ text }) => text)).toEqual(['Alpha ', '', 'Second']);
  expect(paragraphs[0]!.paraId).toBe(id);
  expect(paragraphs[2]!.paraId).toBe(secondId);
  expect(new Set(paragraphs.map(({ paraId }) => paraId)).size).toBe(3);
  expectCaret(broken, { story: 'body', paraId: paragraphs[1]!.paraId, offset: 0 });
  expect(arm.version).toBe(5);
  const beforeSelect = arm.engine.encodeStateVector();
  expect(execute(arm, command(arm, 'select', { selection: selection(id, 2) })).status).toBe('noop');
  expect(arm.engine.encodeStateVector()).toEqual(beforeSelect);
  expect(arm.version).toBe(5);
}, TIMEOUT);

test('line break keeps typed text in one paragraph and saves a wrapping break', async () => {
  await preloadParseWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/parse/docx_parse_bg.wasm'
  ))));
  const arm = await openCommandEngine(['']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'insertText', { text: 'a', selection: selection(id, 0) })).status).toBe('applied');
  const broken = execute(arm, command(arm, 'lineBreak', {}));
  expect(broken.status).toBe('applied');
  expectCaret(broken, { story: 'body', paraId: id, offset: 2 });
  expect(execute(arm, command(arm, 'insertText', { text: 'b' })).status).toBe('applied');
  expect(arm.engine.paragraphs('body').map(paragraph => paragraph.paraId)).toEqual([id]);
  const projected = yrsToDocument(wrapSession(arm.native, 8101), arm.engine.propertyDocument());
  const saved = await repackDocx(projected);
  const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(saved))['word/document.xml']);
  expect(xml.match(/<w:p[\s>]/g)).toHaveLength(1);
  expect(xml).toMatch(/<w:t[^>]*>a<\/w:t>[\s\S]*<w:br(?: w:type="textWrapping")?\s*\/>[\s\S]*<w:t[^>]*>b<\/w:t>/);
}, TIMEOUT);

test('line break replaces a selection and undo restores it in one step', async () => {
  const arm = await openCommandEngine(['abcd']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const broken = execute(arm, command(arm, 'lineBreak', { selection: selection(id, 3, 1) }));
  expect(broken.status).toBe('applied');
  expectCaret(broken, { story: 'body', paraId: id, offset: 2 });
  expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'text')
    .map(segment => segment.text)).toEqual(['a', 'd']);
  expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'embed')
    .map(segment => segment.embedKind)).toEqual(['break']);
  expect(execute(arm, command(arm, 'undo', {})).status).toBe('applied');
  expect(texts(arm)).toEqual(['abcd']);
  expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'embed')).toHaveLength(0);
  expect(execute(arm, command(arm, 'redo', {})).status).toBe('applied');
  expect(arm.engine.paragraphs('body').map(paragraph => paragraph.paraId)).toEqual([id]);
  expect(arm.engine.storySegments('body').filter(segment => segment.kind === 'embed')
    .map(segment => segment.embedKind)).toEqual(['break']);
}, TIMEOUT);

test('d1a_rebases_sticky_anchors_and_predecessor_selection', async () => {
  const arm = await openCommandEngine();
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const selected = execute(arm, command(arm, 'select', { selection: selection(id, 3) }));
  expect(selected.status).toBe('noop');
  const sticky = expectCaret(selected, { story: 'body', paraId: id, offset: 3 });
  expect(sticky.anchor.sticky).toBeInstanceOf(Uint8Array);
  expect(sticky.focus.sticky).toBeInstanceOf(Uint8Array);
  expect(sticky.anchor.sticky!.byteLength).toBeGreaterThan(0);
  const oldSelection = structuredClone(sticky);
  expect(execute(arm, command(arm, 'insertText', { text: '--', selection: selection(id, 0) })).status).toBe('applied');
  const rebased = command(arm, 'insertText', { text: '!', selection: oldSelection }, { basis: { epoch: 1, version: 0 } });
  const resolved = resolveCommandSelection(arm.engine, arm.state, rebased, context(arm, rebased));
  expect(resolved).toMatchObject({ ok: true, selection: {
    anchor: { story: 'body', paraId: id, offset: 5 }, focus: { story: 'body', paraId: id, offset: 5 },
  } });
  const first = execute(arm, rebased);
  expect(first.status).toBe('applied');
  const predecessor = expectCaret(first, { story: 'body', paraId: id, offset: 6 });
  expect(texts(arm)).toEqual(['--Alp!ha beta', 'Second']);
  expect(execute(arm, command(arm, 'select', { selection: selection(id, 0) })).status).toBe('noop');

  const dependent = command(arm, 'insertText', { text: '?' }, {
    basis: { epoch: 1, version: 0 }, afterSeq: rebased.seq,
  });
  const predecessorResolved = resolveCommandSelection(arm.engine, arm.state, dependent, context(arm, dependent, predecessor));
  expect(predecessorResolved).toMatchObject({ ok: true, selection: {
    anchor: { story: 'body', paraId: id, offset: 6 }, focus: { story: 'body', paraId: id, offset: 6 },
  } });
  const second = execute(arm, dependent, predecessor);
  expect(second.status).toBe('applied');
  expect(texts(arm)).toEqual(['--Alp!?ha beta', 'Second']);
  expectCaret(second, { story: 'body', paraId: id, offset: 7 });
  expect(arm.version).toBe(3);
}, TIMEOUT);

test('d1a_refuses_invalid_basis_anchors_and_readonly_before_mutation', async () => {
  const arm = await openCommandEngine();
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  expect(execute(arm, command(arm, 'insertText', { text: 'X' })).status).toBe('applied');
  const forged = selection(id, 4);
  forged.anchor.sticky = arm.engine.encodeStickyPosition(forged.anchor).encoded;
  forged.focus.sticky = forged.anchor.sticky.slice();
  const invalid: DocxCommandEnvelope[] = [
    command(arm, 'insertText', { text: '!', selection: selection(id, 1) }, { basis: { epoch: 1, version: 0 } }),
    command(arm, 'insertText', { text: '!', selection: selection(id, 1) }, { basis: { epoch: 1, version: 2 } }),
    command(arm, 'insertText', { text: '!', selection: forged }, { basis: { epoch: 1, version: 0 } }),
    command(arm, 'insertText', { text: '!', selection: selection(id, -1) }),
    command(arm, 'paragraphBreak', { selection: selection(id, 0, 1000) }),
    command(arm, 'lineBreak', { selection: selection(id, 0, 1000) }),
    command(arm, 'deleteBackward', { selection: selection('missing', 0) }),
    command(arm, 'insertText', { text: '!', selection: selection(id, 0, 0, 'missing-story') }),
    command(arm, 'replaceSelection', { text: '!', selection: {
      anchor: { story: 'body', paraId: id, offset: 0 },
      focus: { story: 'missing-story', paraId: id, offset: 1 },
    } }),
    command(arm, 'formatRange', { selection: selection(id, 0, 1), delta: { fontSize: -1 } }),
    { ...command(arm, 'insertText', { text: '!' }), payload: { text: 42 } } as unknown as DocxCommandEnvelope,
  ];
  for (const [index, intent] of invalid.entries()) {
    const before = snapshot(arm);
    const checked = preflightDocxCommand(arm.engine, arm.state, intent, context(arm, intent));
    expect(checked.status).toBe('refused');
    if (index === 0) expect(checked).toMatchObject({ reason: 'stale-anchor' });
    expect(checked).toHaveProperty('selection');
    expect(checked).toHaveProperty('editState');
    expect(snapshot(arm)).toEqual(before);
  }
  arm.state.storedMarks = { story: 'body', paraId: id, clear: false, delta: { fontSize: -1 } };
  const beforeCompound = snapshot(arm);
  expect(execute(arm, command(arm, 'insertText', { text: 'unsafe' })).status).toBe('refused');
  expect(execute(arm, command(arm, 'lineBreak', {})).status).toBe('refused');
  expect(snapshot(arm)).toEqual(beforeCompound);
  arm.state.storedMarks = null;

  arm.state.readOnly = true;
  for (const intent of [
    command(arm, 'insertText', { text: 'blocked' }),
    command(arm, 'paragraphBreak', { selection: selection(id, 0, 2) }),
    command(arm, 'lineBreak', { selection: selection(id, 0, 2) }),
    command(arm, 'formatRange', { selection: selection(id, 0, 1), delta: { bold: true } }),
  ]) {
    const before = snapshot(arm);
    expect(execute(arm, intent)).toMatchObject({ status: 'refused', editState: { readOnly: true } });
    expect(snapshot(arm)).toEqual(before);
  }
  expect(execute(arm, command(arm, 'setReadOnly', { readOnly: false })).status).toBe('noop');
  expect(arm.state.readOnly).toBe(false);
  expect(execute(arm, command(arm, 'setReadOnly', { readOnly: true })).status).toBe('noop');
  const beforeOrdered = snapshot(arm);
  expect(execute(arm, command(arm, 'deleteForward', {})).status).toBe('refused');
  expect(snapshot(arm)).toEqual(beforeOrdered);
}, TIMEOUT);

test('d1a_refuses_anchorless_stale_commands_before_mutation', async () => {
  const arm = await openCommandEngine(['abc']);
  expect(execute(arm, command(arm, 'insertText', { text: 'X' })).status).toBe('applied');
  expect(arm.state.selection?.focus.sticky).toBeInstanceOf(Uint8Array);
  const extra = { basis: { epoch: 1, version: 0 } };
  for (const intent of [
    command(arm, 'insertText', { text: '!' }, extra),
    command(arm, 'deleteBackward', {}, extra),
    command(arm, 'paragraphBreak', {}, extra),
    command(arm, 'lineBreak', {}, extra),
    command(arm, 'toggleMark', { mark: { type: 'bold' } }, extra),
  ]) {
    const before = snapshot(arm);
    expect(execute(arm, intent)).toMatchObject({ status: 'refused', reason: 'stale-anchor' });
    expect(snapshot(arm)).toEqual(before);
  }
}, TIMEOUT);

test('d1a_unicode_delete_and_paragraph_merge_return_selection', async () => {
  for (const [text, kind, offset, expected, caret] of [
    ['A😀B', 'deleteBackward', 3, 'AB', 1],
    ['A😀B', 'deleteForward', 1, 'AB', 1],
    ['Ae\u0301B', 'deleteBackward', 3, 'AeB', 2],
    ['Ae\u0301B', 'deleteForward', 2, 'AeB', 2],
    ['Ae\u0301B', 'deleteForward', 1, 'A\u0301B', 1],
  ] as const) {
    const arm = await openCommandEngine([text]);
    const id = arm.engine.paragraphs('body')[0]!.paraId;
    const result = execute(arm, command(arm, kind, { selection: selection(id, offset) }));
    expect(result.status).toBe('applied');
    expect(texts(arm)).toEqual([expected]);
    expectCaret(result, { story: 'body', paraId: id, offset: caret });
    expect(arm.version).toBe(1);
  }
  const selectedArm = await openCommandEngine(['A😀B']);
  const selectedId = selectedArm.engine.paragraphs('body')[0]!.paraId;
  const selectedDelete = execute(selectedArm, command(selectedArm, 'deleteBackward', {
    selection: selection(selectedId, 1, 3),
  }));
  expect(selectedDelete.status).toBe('applied');
  expect(texts(selectedArm)).toEqual(['AB']);
  expectCaret(selectedDelete, { story: 'body', paraId: selectedId, offset: 1 });
  expect(selectedArm.version).toBe(1);

  const countedArm = await openCommandEngine(['A😀e\u0301B']);
  const countedId = countedArm.engine.paragraphs('body')[0]!.paraId;
  const countedDelete = execute(countedArm, command(countedArm, 'deleteBackward', {
    selection: selection(countedId, 5), count: 2,
  }));
  expect(countedDelete.status).toBe('applied');
  expect(texts(countedArm)).toEqual(['A😀B']);
  expectCaret(countedDelete, { story: 'body', paraId: countedId, offset: 3 });
  expect(countedArm.version).toBe(1);

  const arm = await openCommandEngine(['A😀', 'e\u0301B']);
  const [first, second] = arm.engine.paragraphs('body');
  const merged = execute(arm, command(arm, 'deleteBackward', { selection: selection(second!.paraId, 0) }));
  expect(merged.status).toBe('applied');
  expect(texts(arm)).toEqual(['A😀e\u0301B']);
  expect(arm.engine.paragraphs('body').map(({ paraId }) => paraId)).toEqual([first!.paraId]);
  expectCaret(merged, { story: 'body', paraId: first!.paraId, offset: 3 });
  expect(arm.version).toBe(1);

  const crossed = await openCommandEngine(['A😀', '', 'B']);
  const crossedParagraphs = crossed.engine.paragraphs('body');
  const crossedDelete = execute(crossed, command(crossed, 'deleteBackward', {
    selection: selection(crossedParagraphs[2]!.paraId, 1), count: 4,
  }));
  expect(crossedDelete.status).toBe('applied');
  expect(texts(crossed)).toEqual(['A']);
  expectCaret(crossedDelete, { story: 'body', paraId: crossedParagraphs[0]!.paraId, offset: 1 });
}, TIMEOUT);

test('d1a_formatting_and_stored_marks_apply_in_worker', async () => {
  const arm = await openCommandEngine(['abc']);
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const all = selection(id, 0, 3);
  expect(execute(arm, command(arm, 'toggleMark', { selection: all, mark: { type: 'bold' } })).status).toBe('applied');
  expect(arm.engine.selectionContext(storyRange(all)).bold).toBe(true);
  expect(execute(arm, command(arm, 'formatRange', {
    selection: all, delta: { italic: true, fontSize: 18, fontFamily: { ascii: 'Arial' }, color: { rgb: '123456' } },
  })).status).toBe('applied');
  expect(arm.engine.selectionContext(storyRange(all))).toMatchObject({
    bold: true, italic: true, fontSize: 36, fontFamily: 'Arial', color: '123456',
  });
  expect(execute(arm, command(arm, 'clearFormatting', { selection: all })).status).toBe('applied');
  expect(arm.engine.selectionContext(storyRange(all))).toMatchObject({ bold: false, italic: false, fontSize: null });
  expect(execute(arm, command(arm, 'setParagraphAttrs', { selection: all, attrs: { alignment: 'center' } })).status).toBe('applied');
  expect(arm.engine.selectionContext(storyRange(all)).alignment).toBe('center');
  expect(execute(arm, command(arm, 'applyParagraphStyle', { selection: all, styleId: 'Heading1' })).status).toBe('applied');
  expect(arm.engine.selectionContext(storyRange(all)).styleId).toBe('Heading1');
  expect(arm.version).toBe(5);

  const stored = await openCommandEngine(['abc']);
  const storedId = stored.engine.paragraphs('body')[0]!.paraId;
  const before = snapshot(stored);
  const toggled = execute(stored, command(stored, 'toggleMark', { mark: { type: 'bold' } }));
  expect(toggled.status).toBe('noop');
  expect(toggled.editState).toMatchObject({ storedMarks: { delta: { bold: true } }, canUndo: false, canRedo: false });
  expect(snapshot(stored)).toMatchObject({ texts: before.texts, ids: before.ids, version: before.version, vector: before.vector, canUndo: false });
  const inserted = execute(stored, command(stored, 'insertText', { text: 'X' }));
  expect(inserted.status).toBe('applied');
  expect(texts(stored)).toEqual(['Xabc']);
  expectCaret(inserted, { story: 'body', paraId: storedId, offset: 1 });
  expect(stored.engine.selectionContext(storyRange(selection(storedId, 0, 1))).bold).toBe(true);
  expect(stored.version).toBe(1);
  expect(execute(stored, command(stored, 'clearFormatting', {})).status).toBe('noop');
  expect(stored.state.storedMarks?.clear).toBe(true);
  expect(execute(stored, command(stored, 'insertText', { text: 'Y' })).status).toBe('applied');
  expect(texts(stored)).toEqual(['XYabc']);
  expect(stored.engine.selectionContext(storyRange(selection(storedId, 1, 2))).bold).toBe(false);
  expect(stored.version).toBe(2);
}, TIMEOUT);

test('d1a_rebased_caret_keeps_stored_marks_after_remote_insertion', async () => {
  const arm = await openCommandEngine(['abcdef'], { clientId: 8211 });
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  execute(arm, command(arm, 'select', { selection: selection(id, 3) }));
  const toggled = execute(arm, command(arm, 'toggleMark', { mark: { type: 'bold' } }));
  const sticky = expectCaret(toggled, { story: 'body', paraId: id, offset: 3 });
  const remote = await openCommandEngine(['abcdef'], { clientId: 8212, checkpoint: arm.engine.encodeState() });
  const vector = arm.engine.encodeStateVector();
  remote.engine.insertText({ story: 'body', paraId: id, offset: 0 }, '--');
  arm.engine.applyUpdate(remote.engine.encodeStateAsUpdate(vector));
  expect(arm.state.selection?.focus.offset).toBe(3);
  const inserted = execute(arm, command(arm, 'insertText', { text: 'X', selection: sticky }));
  expect(inserted.status).toBe('applied');
  expectCaret(inserted, { story: 'body', paraId: id, offset: 6 });
  expect(texts(arm)).toEqual(['--abcXdef']);
  expect(arm.engine.selectionContext(storyRange(selection(id, 5, 6))).bold).toBe(true);
  expect(inserted.editState?.storedMarks?.delta.bold).toBe(true);
}, TIMEOUT);

test('d1a_command_host_uses_target_reads_and_one_state_projection', async () => {
  const arm = await openCommandEngine(['abc', ...Array.from({ length: 80 }, () => 'untouched')]);
  const init = {
    session: 'd1-worker', incarnation: 1, epoch: 1, readOnly: false,
    fonts: [], layoutInput: '{}', layoutWithRegions: true, layoutRevision: 1, extras: '{}',
  };
  const hooks = createResidentDocxCommandHooks(arm.engine, init);
  const host = createResidentCommandHost({
    init, registry: { handlers: {}, policies: {} }, hooks: {
      ...hooks,
      frame: ({ coversThroughSeq }) => ({ coversThroughSeq, revision: coversThroughSeq + 1, payload: {
        frame: new ArrayBuffer(0), caret: { frameEpoch: coversThroughSeq + 1, caretRect: null }, layoutRevision: 1,
      } }),
      snapshot: () => arm.engine.encodeState(),
    },
  });
  await host.initial();
  const paragraphs = spyOn(arm.engine, 'paragraphs');
  const story = spyOn(arm.engine, 'storySegments');
  const targets = spyOn(arm.engine, 'paragraphSegments');
  const toolbar = spyOn(arm.engine, 'selectionContext');
  const locations = spyOn(arm.engine.geometryReader, 'locateParagraph');
  try {
    const base = { session: init.session, incarnation: 1 };
    const intents: DocxCommandIntent[] = [
      { ...base, seq: 1, basis: { epoch: 1, version: 0 }, kind: 'insertText', payload: { text: 'X' } },
      { ...base, seq: 2, basis: { epoch: 1, version: 1 }, kind: 'deleteBackward', payload: {} },
      { ...base, seq: 3, basis: { epoch: 1, version: 2 }, kind: 'toggleMark', payload: { mark: { type: 'bold' } } },
      { ...base, seq: 4, basis: { epoch: 1, version: 0 }, kind: 'insertText', payload: { text: 'blocked' } },
      { ...base, seq: 5, basis: { epoch: 1, version: 2 }, kind: 'setReadOnly', payload: { readOnly: true } },
    ];
    for (const intent of intents) {
      const { seq } = intent;
      targets.mockClear(); toolbar.mockClear(); locations.mockClear();
      host.policy(intent);
      const reply = await host.handler(undefined, intent);
      const ack = isTransferResult(reply) ? reply.value : reply;
      expect(ack.outcomes[0]?.status).toBe(seq === 4 ? 'refused' : seq >= 3 ? 'noop' : 'applied');
      expect(toolbar).toHaveBeenCalledTimes(1);
      expect(targets.mock.calls.every(([story, indices]) => story === 'body' && indices.length === 1 && indices[0] === 0)).toBe(true);
      expect(targets.mock.calls.length).toBeLessThanOrEqual(2);
      expect(locations.mock.calls.length).toBeLessThan(20);
      expect(paragraphs).not.toHaveBeenCalled();
      expect(story).not.toHaveBeenCalled();
    }
  } finally {
    locations.mockRestore(); toolbar.mockRestore(); targets.mockRestore(); story.mockRestore(); paragraphs.mockRestore();
  }
}, TIMEOUT);

test('d1a_postmutation_failure_is_unknown_outcome_without_retry', async () => {
  const arm = await openCommandEngine(['abc']);
  const intent = command(arm, 'insertText', { text: 'X' });
  const ctx = context(arm, intent);
  const checked = preflightDocxCommand(arm.engine, arm.state, intent, ctx);
  expect(checked.status).toBe('ready');
  if (checked.status !== 'ready') throw new Error('Valid insert was refused');
  const afterMutation = <A extends unknown[], R>(mutate: (...args: A) => R) => (...args: A): never => {
    mutate(...args);
    throw new Error('receipt failed after insertion');
  };
  const insert = arm.native.insert_text.bind(arm.native);
  const replace = arm.native.replace_range.bind(arm.native);
  const input = arm.native.apply_input.bind(arm.native);
  const inputProfiled = arm.native.apply_input_profiled.bind(arm.native);
  const resident = arm.native.edit_text_resident.bind(arm.native);
  const inserted = spyOn(arm.native, 'insert_text').mockImplementation(afterMutation(insert));
  const residentEdited = spyOn(arm.native, 'edit_text_resident').mockImplementation(afterMutation(resident));
  const replaced = spyOn(arm.native, 'replace_range').mockImplementation(afterMutation(replace));
  const fast = spyOn(arm.native, 'apply_input').mockImplementation(afterMutation(input));
  const fastProfiled = spyOn(arm.native, 'apply_input_profiled').mockImplementation(afterMutation(inputProfiled));
  const host = createCommandHost<DocxCommandPayload, DocxCommandSelection | null, DocxEditState, null>({
    session: intent.session, incarnation: intent.incarnation, epoch: 1,
    getVersion: () => arm.version,
    apply: (admitted, applyContext) => {
      const preflight = preflightDocxCommand(arm.engine, arm.state, admitted, applyContext);
      if (preflight.status === 'refused') return preflight;
      const result = applyDocxCommand(arm.engine, arm.state, preflight.prepared, applyContext);
      if (result.status === 'applied') arm.version += 1;
      return result;
    },
    frame: ({ coversThroughSeq }) => ({ coversThroughSeq, revision: 1, payload: null }),
    snapshot: () => arm.engine.encodeState(),
  });
  let failure: unknown;
  try {
    host.policy(intent);
    try { await host.handler(undefined, intent); }
    catch (error) { failure = error; }
    expect(isCommandOutcomeUnknownError(failure)).toBe(true);
    expect(failure).toMatchObject({ name: 'CommandOutcomeUnknown', refusal: { seq: intent.seq } });
    expect(inserted.mock.calls.length + residentEdited.mock.calls.length + replaced.mock.calls.length +
      fast.mock.calls.length + fastProfiled.mock.calls.length).toBe(1);
    expect(texts(arm)).toEqual(['Xabc']);
    expect(arm.native.can_undo()).toBe(true);
    expect(arm.version).toBe(0);
    const later = command(arm, 'insertText', { text: 'retry' });
    host.policy(later);
    await expect(host.handler(undefined, later)).rejects.toMatchObject({
      name: 'CommandOutcomeUnknown', refusal: { seq: intent.seq },
    });
    expect(inserted.mock.calls.length + residentEdited.mock.calls.length + replaced.mock.calls.length +
      fast.mock.calls.length + fastProfiled.mock.calls.length).toBe(1);
    expect(texts(arm)).toEqual(['Xabc']);
  } finally {
    fastProfiled.mockRestore(); fast.mockRestore(); replaced.mockRestore(); residentEdited.mockRestore(); inserted.mockRestore();
  }
}, TIMEOUT);

test.todo('d1a_live_continuation_matches_restored_replay: DX14 needs native allocator checkpoint persistence', async () => {
  const left = await openCommandEngine(['abcd'], { clientId: 8191 });
  expect(execute(left, command(left, 'insertText', {
    text: 'Q', suggesting: { name: 'D1 author', date: '2026-01-01T00:00:00Z' },
  })).status).toBe('applied');
  const checkpoint = left.engine.encodeState();
  const right = await openCommandEngine(['abcd'], { clientId: 8191, checkpoint });
  right.version = left.version;
  right.seq = left.seq;
  const id = left.engine.paragraphs('body')[0]!.paraId;
  expect(right.engine.paragraphs('body')[0]!.paraId).toBe(id);
  const intents: DocxCommandEnvelope[] = [
    command(left, 'paragraphBreak', { selection: selection(id, 2) }, { replay: true }),
    command(left, 'insertText', { text: 'X' }, { basis: { epoch: 1, version: 1 }, afterSeq: 2, replay: true }),
    command(left, 'paragraphBreak', {}, { basis: { epoch: 1, version: 1 }, afterSeq: 3, replay: true }),
  ];
  const replay = (arm: typeof left) => {
    let predecessor: DocxCommandSelection | null | undefined;
    const selections: DocxCommandSelection[] = [];
    for (const intent of intents) {
      const result = execute(arm, intent, predecessor);
      expect(result.status).toBe('applied');
      expect(result.selection).toBeDefined();
      if (!result.selection) throw new Error('Replay result lacks selection');
      expect(result.selection.anchor.sticky).toBeInstanceOf(Uint8Array);
      expect(result.selection.focus.sticky).toBeInstanceOf(Uint8Array);
      selections.push(structuredClone(result.selection));
      predecessor = result.selection;
    }
    return selections;
  };
  expect(replay(left)).toEqual(replay(right));
  expect(texts(left)).toEqual(['Qa', 'X', 'bcd']);
  expect(texts(right)).toEqual(texts(left));
  expect(right.engine.paragraphIdentities()).toEqual(left.engine.paragraphIdentities());
  expect(right.version).toBe(4);
  expect(left.version).toBe(4);
}, TIMEOUT);

test('setObjectProperties updates image wrapType, float/position attrs, and transform and round-trips to OOXML', async () => {
  await preloadParseWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/parse/docx_parse_bg.wasm'
  ))));
  const bodyXml = '<w:p w14:paraId="00000001" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="7" name="Picture" descr="First"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLink"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
    '<w:r><w:t>Seed</w:t></w:r></w:p>';
  const arm = await openCommandEngine(undefined, { bodyXml });
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const initial = readResidentProperties(arm.engine, { kind: 'objectSelection', selection: selection(id, 0, 1) });
  expect(initial).toMatchObject({ kind: 'object', properties: { wrapType: 'inline', displayMode: 'inline' } });
  if (initial?.kind !== 'object') throw new Error('Expected image object properties');
  const { handle } = initial;

  const anchorPosition = {
    horizontal: { relativeTo: 'column', posOffset: 0 },
    vertical: { relativeTo: 'paragraph', posOffset: 0 },
  };
  const savedDocumentXml = async () => {
    const projected = yrsToDocument(wrapSession(arm.native, 8101), arm.engine.propertyDocument());
    const saved = await repackDocx(projected);
    return new TextDecoder().decode(unzipContainer(new Uint8Array(saved))['word/document.xml']);
  };

  for (const wrap of ['square', 'tight', 'through', 'topAndBottom', 'inFront', 'behind'] as const) {
    const displayMode = wrap === 'topAndBottom' ? 'block' : 'float';
    const cssFloat = wrap === 'square' || wrap === 'tight' || wrap === 'through' ? 'left' : 'none';
    const wrapText = cssFloat === 'left' ? 'right' : 'bothSides';
    const applied = execute(arm, command(arm, 'setObjectProperties', {
      handle,
      expectVersion: arm.engine.proposalEngine.version(),
      patch: {
        wrap,
        wrapType: wrap,
        displayMode,
        cssFloat,
        wrapText,
        position: anchorPosition,
        transform: 'rotate(90deg) scaleX(-1) scaleY(-1)',
      },
    }));
    expect(applied.status).toBe('applied');
    const current = readResidentProperties(arm.engine, { kind: 'objectSelection', selection: selection(id, 0, 1) });
    expect(current).toMatchObject({
      kind: 'object',
      handle,
      properties: {
        wrapType: wrap,
        displayMode,
        cssFloat,
        wrapText,
        transform: 'rotate(90deg) scaleX(-1) scaleY(-1)',
      },
    });
    if (wrap === 'square') {
      const xml = await savedDocumentXml();
      expect(xml).toContain('<wp:anchor');
      expect(xml).toContain('<wp:wrapSquare');
      expect(xml).toContain('rot="5400000"');
      expect(xml).toContain('flipH="1"');
      expect(xml).toContain('flipV="1"');
    }
    if (wrap === 'behind') {
      const xml = await savedDocumentXml();
      expect(xml).toContain('<wp:anchor');
      expect(xml).toContain('behindDoc="1"');
      expect(xml).toContain('<wp:wrapNone/>');
      expect(xml).toContain('rot="5400000"');
      expect(xml).toContain('flipH="1"');
      expect(xml).toContain('flipV="1"');
    }
  }

  const reverted = execute(arm, command(arm, 'setObjectProperties', {
    handle,
    expectVersion: arm.engine.proposalEngine.version(),
    patch: {
      wrap: 'inline',
      wrapType: 'inline',
      displayMode: 'inline',
      cssFloat: 'none',
      wrapText: null,
      position: null,
      distTop: null,
      distBottom: null,
      distLeft: null,
      distRight: null,
      transform: null,
    },
  }));
  expect(reverted.status).toBe('applied');
  const finalProps = readResidentProperties(arm.engine, { kind: 'objectSelection', selection: selection(id, 0, 1) });
  expect(finalProps).toMatchObject({
    kind: 'object',
    handle,
    properties: { wrapType: 'inline', displayMode: 'inline', cssFloat: 'none' },
  });
  if (finalProps?.kind === 'object') {
    expect(finalProps.properties.transform).toBeUndefined();
    expect(finalProps.properties.position).toBeUndefined();
  }
  const inlineXml = await savedDocumentXml();
  expect(inlineXml).toContain('<wp:inline');
}, TIMEOUT);

test('setObjectProperties preserves image wrap distances across non-inline worker wrap patches', async () => {
  await preloadParseWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/parse/docx_parse_bg.wasm'
  ))));
  const bodyXml = '<w:p w14:paraId="00000001" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<w:r><w:drawing><wp:inline distT="114300" distB="228600" distL="342900" distR="457200"><wp:extent cx="952500" cy="476250"/><wp:docPr id="7" name="Picture" descr="First"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLink"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
    '<w:r><w:t>Seed</w:t></w:r></w:p>';
  const arm = await openCommandEngine(undefined, { bodyXml });
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const initial = readResidentProperties(arm.engine, { kind: 'objectSelection', selection: selection(id, 0, 1) });
  if (initial?.kind !== 'object') throw new Error('Expected image object properties');
  const { handle } = initial;
  const buildWorkerWrapPatch = (currentAttrs: Readonly<Record<string, unknown>>, target: ImageLayoutTarget): Record<string, unknown> => {
    const next = resolveImageLayoutAttrs(target, currentAttrs as ImageLayoutCurrentAttrs);
    return {
      wrap: next.wrapType,
      wrapType: next.wrapType,
      displayMode: next.displayMode ?? null,
      cssFloat: next.cssFloat ?? null,
      wrapText: next.wrapText ?? null,
      position: next.position ?? null,
      ...('distTop' in next ? { distTop: next.distTop ?? null } : {}),
      ...('distBottom' in next ? { distBottom: next.distBottom ?? null } : {}),
      ...('distLeft' in next ? { distLeft: next.distLeft ?? null } : {}),
      ...('distRight' in next ? { distRight: next.distRight ?? null } : {}),
    };
  };
  const savedDocumentXml = async () => {
    const projected = yrsToDocument(wrapSession(arm.native, 8101), arm.engine.propertyDocument());
    const saved = await repackDocx(projected);
    return new TextDecoder().decode(unzipContainer(new Uint8Array(saved))['word/document.xml']);
  };
  for (const target of ['squareLeft', 'behind'] as const) {
    const snapshot = readResidentProperties(arm.engine, { kind: 'objectSelection', selection: selection(id, 0, 1) });
    if (snapshot?.kind !== 'object') throw new Error('Expected image object properties');
    const applied = execute(arm, command(arm, 'setObjectProperties', {
      handle,
      expectVersion: arm.engine.proposalEngine.version(),
      patch: buildWorkerWrapPatch(snapshot.properties, target),
    }));
    expect(applied.status).toBe('applied');
    const xml = await savedDocumentXml();
    expect(xml).toContain('<wp:anchor');
    expect(xml).toContain('distT="114300"');
    expect(xml).toContain('distB="228600"');
    expect(xml).toContain('distL="342900"');
    expect(xml).toContain('distR="457200"');
  }
}, TIMEOUT);

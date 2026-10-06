import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { SessionFailure } from '../../../../shared/office-session/types';
import { initWasm, openPresentation } from '../wasm/loader';
import { PptxDocument } from '../wasm/generated/pptx_wasm.js';
import { createPresentationEditorSession } from './editorSession';
import { createPresentationSessionHost } from './host';
import type { PptxWorkerEditorOperation } from './replay';

let source: Uint8Array;
let module: WebAssembly.Module;
let font: Uint8Array;
beforeAll(async () => {
  const root = resolve(import.meta.dir, '../../../..');
  const [wasm, deck, face] = await Promise.all([
    readFile(resolve(import.meta.dir, '../wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  module = await WebAssembly.compile(wasm);
  await initWasm(module);
  source = deck;
  font = face;
});

async function setup() {
  const pair = createInProcessPair();
  const calls: { method: string; args: unknown[] }[] = [];
  const held: (() => void)[] = [];
  const controls = { holdRequests: false, holdReplies: false, postFailure: false,
    mismatch: '' as string, refusal: false };
  let crash!: (cause: unknown) => void;
  let replayObserved!: () => void;
  const observed = new Promise<void>((resolve) => { replayObserved = resolve; });
  let transportClosed!: () => void;
  const closed = new Promise<void>((resolve) => { transportClosed = resolve; });
  const workerOpen = spyOn(PptxDocument, 'openReplayBaseline');
  const peerOpen = spyOn(PptxDocument, 'openPeerDeckJson');
  createPresentationSessionHost({ ...pair.host, close() { pair.host.close(); transportClosed(); }, post(message, transfer) {
    if (isHostMessage(message) && message.kind === 'reply' && message.ok &&
      typeof message.value === 'object' && message.value !== null && 'consumed' in message.value) {
      replayObserved();
      let reply = message;
      if (controls.mismatch && controls.mismatch !== 'expectedOutcome') {
        const value = reply.value as Record<string, unknown>;
        const field = controls.mismatch;
        reply = { ...reply, value: { ...value, [field]: field === 'outcome'
          ? { ...(value.outcome as object), applied: 'wrong' } : 'wrong' } };
      }
      message = controls.refusal
        ? { ...reply, ok: false, error: { name: 'Error', message: 'refused' } } : reply;
      if (controls.holdReplies) { held.push(() => pair.host.post(message, transfer)); return; }
    }
    pair.host.post(message, transfer);
  } }, { initEditorWasm: async () => module });
  let errors = 0;
  const faces = [{ family: 'Liberation Sans', bytes: font }];
  const owner = createPresentationEditorSession(source, { wasm: module, clientId: 7201, fonts: faces,
    fallbackFonts: faces, onError: () => { errors += 1; } }, {
    ...pair.client,
    onError(listener) { crash = listener; return pair.client.onError(listener); },
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call') {
        calls.push({ method: message.method, args: structuredClone(message.args) });
        if (message.method === 'replay') {
          if (controls.postFailure) throw new Error('replay post failed');
          if (controls.holdRequests) { replayObserved(); held.push(() => pair.client.post(message, transfer)); return; }
          if (controls.mismatch === 'expectedOutcome') {
            const envelope = message.args[0] as { expectedOutcome: object };
            message = { ...message, args: [{ ...envelope,
              expectedOutcome: { ...envelope.expectedOutcome, applied: false } }] };
          }
        }
      }
      pair.client.post(message, transfer);
    },
  });
  await owner.whenHydrated();
  const workerDoc = workerOpen.mock.results[0]!.value as PptxDocument;
  const peerDoc = peerOpen.mock.results[0]!.value as PptxDocument;
  workerOpen.mockRestore();
  peerOpen.mockRestore();
  const access = await owner.handleAsync();
  return { owner, access, workerDoc, peerDoc, calls, controls, observed, closed,
    get errors() { return errors; },
    crash: () => crash(new SessionFailure('crash', 'worker lost')),
    hostDispose: () => pair.client.post({ protocol: 1, kind: 'dispose' }),
    release: () => { for (const send of held.splice(0)) send(); },
  };
}

async function equal(context: Awaited<ReturnType<typeof setup>>) {
  const { owner, access, workerDoc, peerDoc } = context;
  await owner.flush();
  expect(peerDoc.documentVersion()).toBe(workerDoc.documentVersion());
  expect(peerDoc.clientId).toBe(workerDoc.clientId);
  expect(JSON.parse(peerDoc.snapshotJson())).toEqual(JSON.parse(workerDoc.snapshotJson()));
  expect(JSON.parse(peerDoc.readContentJson('{}'))).toEqual(JSON.parse(workerDoc.readContentJson('{}')));
  expect(JSON.parse(peerDoc.listProposalsJson())).toEqual(JSON.parse(workerDoc.listProposalsJson()));
  expect(peerDoc.canUndo()).toBe(workerDoc.canUndo());
  expect(peerDoc.canRedo()).toBe(workerDoc.canRedo());
  expect(peerDoc.undoCaptureMode()).toBe('manual');
  expect(workerDoc.undoCaptureMode()).toBe('manual');
  expect(peerDoc.saveBytes()).toEqual(await owner.saveAsync());
  for (const [index, slide] of access.snapshot().slides.entries()) {
    const frame = await owner.frame(slide.id);
    expect(frame.sequence).toBe(owner.state.sequence);
    expect(frame.version).toBe(access.version());
    expect(frame.displayList).toEqual(access.layoutSlide(index));
    for (const shape of slide.shapes) for (const story of shape.textStories) {
      expect(peerDoc.storyJson(JSON.stringify({ storyId: story.id })))
        .toBe(workerDoc.storyJson(JSON.stringify({ storyId: story.id })));
    }
  }
}

async function mixed(context: Awaited<ReturnType<typeof setup>>, parity: boolean) {
  const { owner, access } = context;
  const run = async (op: PptxWorkerEditorOperation) => {
    owner.apply({ method: 'addUndoBoundary', args: [] });
    if (parity) await equal(context);
    const reply = owner.apply(op);
    expect(reply.consumed).toBe(true);
    if (parity) await equal(context);
  };
  await run({ method: 'insertSlide', args: [1] });
  const slide = access.snapshot().slides[1].id;
  await run({ method: 'moveSlide', args: [slide, 0] });
  const rect = { x: 100000, y: 100000, width: 2000000, height: 900000 };
  await run({ method: 'addTextBox', args: [slide, { name: 'Session text', rect, text: 'Hello', style: {} }] });
  const textbox = access.snapshot().slides[0].shapes.find((shape) => shape.name === 'Session text')!;
  const story = textbox.textStories[0].id;
  await run({ method: 'deleteText', args: [story, 0, 5] });
  await run({ method: 'insertText', args: [story, 0, 'Replacement', { bold: true }] });
  await run({ method: 'insertParagraphBreak', args: [story, 2] });
  await run({ method: 'formatText', args: [story, 0, 2, { italic: true }] });
  await run({ method: 'setParagraphAlignment', args: [story, 0, 2, 'ctr'] });
  await run({ method: 'addShape', args: [slide, { name: 'Session shape', geometry: 'roundRect', rect, fill: '#123456' }] });
  const shape = access.snapshot().slides[0].shapes.find((item) => item.name === 'Session shape')!.id;
  const ops: PptxWorkerEditorOperation[] = [
    { method: 'moveShape', args: [slide, shape, 200000, 300000] },
    { method: 'resizeShape', args: [slide, shape, 2100000, 1000000] },
    { method: 'setShapeRect', args: [slide, shape, rect] },
    { method: 'setShapeFill', args: [slide, shape, '#654321'] },
    { method: 'setShapeStroke', args: [slide, shape, { color: '#abcdef', widthPt: 2 }] },
    { method: 'setShapeAdjust', args: [slide, shape, { adj: 25000 }] },
    { method: 'bringShapeToFront', args: [slide, textbox.id] },
    { method: 'sendShapeToBack', args: [slide, textbox.id] },
    { method: 'bringShapeForward', args: [slide, textbox.id] },
    { method: 'sendShapeBackward', args: [slide, textbox.id] },
    { method: 'addPicture', args: [slide, { name: 'Session image', rect, contentType: 'image/png',
      mediaBase64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=' }] },
    { method: 'setCommentFlavor', args: ['modern'] },
    { method: 'addComment', args: [slide, { author: 'Session', text: 'Root', created: '2026-10-06T12:00:00Z' }] },
  ];
  for (const op of ops) await run(op);
  const comment = access.comments().find((item) => item.text === 'Root')!.id;
  await run({ method: 'replyToComment', args: [comment, { author: 'Reply', text: 'Child', created: '2026-10-06T12:01:00Z' }] });
  await run({ method: 'setCommentStatus', args: [comment, true] });
  await run({ method: 'setCommentPosition', args: [comment, { xEmu: 100, yEmu: 200 }] });
  await run({ method: 'removeComment', args: [comment] });
  await run({ method: 'setSlideNotes', args: [slide, 'Notes'] });
  await run({ method: 'propose', args: ['agent', 'Review', [{ type: 'setSlideNotes', slideId: slide, text: 'Accepted' }]] });
  await run({ method: 'acceptProposal', args: [access.listProposals()[0].id] });
  await run({ method: 'propose', args: ['agent', null, [{ type: 'setSlideNotes', slideId: slide, text: 'Rejected' }]] });
  await run({ method: 'rejectProposal', args: [access.listProposals()[0].id] });
  await run({ method: 'applyEdits', args: [{ expectVersion: access.version(), history: 'separate',
    steps: [{ op: 'setSlideNotes', target: { slideId: slide }, text: 'Batch notes' }] }] });
  await run({ method: 'removeShape', args: [slide, shape] });
  await run({ method: 'deleteSlide', args: [slide] });
  owner.apply({ method: 'addUndoBoundary', args: [] });
  if (parity) {
    const final = access.snapshot();
    let undos = 0;
    while (access.canUndo()) {
      expect(++undos).toBeLessThan(100);
      owner.apply({ method: 'undo', args: [] });
      await equal(context);
    }
    let redos = 0;
    while (access.canRedo()) {
      expect(++redos).toBeLessThan(100);
      owner.apply({ method: 'redo', args: [] });
      await equal(context);
    }
    expect(redos).toBe(undos);
    expect(access.snapshot()).toEqual(final);
  }
}

test('mixed_inputs_replay_in_peer_commit_order', async () => {
  const context = await setup();
  context.controls.holdRequests = true;
  try {
    await mixed(context, false);
    await context.observed;
    const finalSequence = context.owner.state.sequence;
    expect(context.owner.state.acknowledgedSequence).toBe(0);
    context.controls.holdRequests = false;
    context.release();
    await context.owner.flush();
    const sent = context.calls.filter((call) => call.method === 'replay');
    expect(sent.map((call) => (call.args[0] as { sequence: number }).sequence))
      .toEqual(Array.from({ length: finalSequence }, (_, index) => index + 1));
    await equal(context);
  } finally { await context.owner.dispose(); }
});

test('every_mixed_edit_prefix_matches_worker_observable_state', async () => {
  const context = await setup();
  try { await equal(context); await mixed(context, true); }
  finally { await context.owner.dispose(); }
}, 120_000);

test('successful_noops_consume_sequence_refusals_do_not', async () => {
  const context = await setup();
  try {
    const reply = context.owner.apply({ method: 'undo', args: [] });
    expect(reply.consumed).toBe(true);
    expect(reply.outcome.applied).toBe(false);
    const refused = context.owner.apply({ method: 'applyEdits', args: [{ expectVersion: 'stale', steps: [] }] });
    expect(refused.consumed).toBe(false);
    expect(context.owner.state.sequence).toBe(1);
    context.access.addUndoBoundary();
    await context.owner.flush();
    expect(context.owner.state.acknowledgedSequence).toBe(2);
    expect(context.owner.state.stage).toBe('ready');
  } finally { await context.owner.dispose(); }
});

test('arguments_are_owned_before_peer_apply', async () => {
  const context = await setup();
  try {
    const draft = { name: 'Owned', rect: { x: 0, y: 0, width: 1000000, height: 1000000 }, text: 'Original', style: {} };
    const off = context.access.onUpdate(() => { draft.text = 'Listener mutation'; draft.rect.x = 900000; });
    context.access.addTextBox(context.access.snapshot().slides[0].id, draft);
    off();
    draft.text = 'Caller mutation';
    await equal(context);
    expect(context.access.snapshot().slides[0].shapes.find((shape) => shape.name === 'Owned')!.x).toBe(0);
    const before = context.owner.state.sequence;
    expect(() => context.owner.apply({ method: 'insertSlide', args: [NaN] })).toThrow();
    expect(context.owner.state.sequence).toBe(before);
    expect(context.owner.state.stage).toBe('ready');
  } finally { await context.owner.dispose(); }
});

test('reentrant_mutation_is_visibly_refused', async () => {
  const context = await setup();
  let refused = false;
  try {
    const slide = context.access.snapshot().slides[0].id;
    const off = context.access.onUpdate(() => {
      expect(context.owner.state.sequence).toBe(1);
      expect(() => context.access.setSlideNotes(slide, 'reentrant')).toThrow('cannot reenter');
      refused = true;
    });
    context.access.setSlideNotes(slide, 'accepted');
    off();
    expect(refused).toBe(true);
    expect(context.owner.state.sequence).toBe(1);
    await equal(context);
  } finally { await context.owner.dispose(); }
});

test('flush_and_save_wait_for_final_ack', async () => {
  const context = await setup();
  const slide = context.access.snapshot().slides[0].id;
  context.controls.holdReplies = true;
  let pending: Promise<PromiseSettledResult<unknown>[]> | undefined;
  try {
    context.access.setSlideNotes(slide, 'before fence');
    let flushed = false;
    let saved = false;
    const flush = context.owner.flush().then(() => { flushed = true; });
    const save = context.owner.saveAsync().then((bytes) => { saved = true; return bytes; });
    pending = Promise.allSettled([flush, save]);
    context.access.setSlideNotes(slide, 'after fence');
    await context.observed;
    expect(flushed).toBe(false);
    expect(saved).toBe(false);
    expect(context.calls.some((call) => call.method === 'editorSave')).toBe(false);
    context.controls.holdReplies = false;
    context.release();
    await flush;
    const reopened = openPresentation(await save);
    try { expect(reopened.snapshot().slides[0].notes).toBe('before fence'); }
    finally { reopened.dispose(); }
    await context.owner.flush();
    expect(context.access.snapshot().slides[0].notes).toBe('after fence');
  } finally {
    await context.owner.dispose();
    await pending;
  }
});

test('normal_save_and_export_use_worker_only', async () => {
  const context = await setup();
  const peerSave = spyOn(context.peerDoc, 'saveBytes');
  const workerSave = spyOn(context.workerDoc, 'saveBytes');
  try {
    context.access.setSlideNotes(context.access.snapshot().slides[0].id, 'worker save');
    await context.owner.saveAsync();
    expect(workerSave).toHaveBeenCalledTimes(1);
    expect(peerSave).not.toHaveBeenCalled();
    expect('exportStructured' in context.access).toBe(false);
    expect('exportMarkdown' in context.access).toBe(false);
    expect('applyUpdate' in context.access).toBe(false);
    expect('setUndoCaptureMode' in context.access).toBe(false);
  } finally { peerSave.mockRestore(); workerSave.mockRestore(); await context.owner.dispose(); }
});

async function recovery(context: Awaited<ReturnType<typeof setup>>, text: string) {
  const result = await context.owner.recoverySave();
  expect(result.recovery).toBe(true);
  const reopened = openPresentation(result.bytes);
  try { expect(reopened.snapshot().slides[0].notes).toBe(text); }
  finally { reopened.dispose(); }
  const retained = result.bytes.slice();
  result.bytes.fill(0);
  expect((await context.owner.recoverySave()).bytes).toEqual(retained);
  await context.owner.dispose();
  expect((await context.owner.recoverySave()).bytes).toEqual(retained);
  expect(() => context.access.snapshot()).toThrow('disposed');
  await expect(context.owner.saveAsync()).rejects.toMatchObject({ code: 'editor-disposed' });
}

test('recovery_survives_host_dispose', async () => {
  const context = await setup();
  try {
    context.access.setSlideNotes(context.access.snapshot().slides[0].id, 'host dispose');
    await context.owner.flush();
    context.hostDispose();
    await context.closed;
    await expect(context.owner.flush()).rejects.toMatchObject({ code: 'editor-failed' });
    await recovery(context, 'host dispose');
  } finally { await context.owner.dispose(); }
});

test('recovery_survives_replay_post_failure', async () => {
  const context = await setup();
  context.controls.postFailure = true;
  try {
    context.access.setSlideNotes(context.access.snapshot().slides[0].id, 'post failure');
    await expect(context.owner.flush()).rejects.toMatchObject({ code: 'editor-failed' });
    await recovery(context, 'post failure');
  } finally { await context.owner.dispose(); }
});

test('recovery_survives_outcome_mismatch', async () => {
  for (const field of ['sequence', 'revision', 'version', 'engineVersion', 'consumed', 'outcome', 'refusal', 'expectedOutcome']) {
    const context = await setup();
    context.controls.mismatch = field === 'refusal' ? '' : field;
    context.controls.refusal = field === 'refusal';
    try {
      context.access.setSlideNotes(context.access.snapshot().slides[0].id, field);
      await expect(context.owner.flush()).rejects.toMatchObject({ code: 'editor-failed' });
      if (field === 'expectedOutcome') expect(context.owner.failure?.cause).toMatchObject({ code: 'outcomeMismatch' });
      expect(context.errors).toBe(1);
      await recovery(context, field);
    } finally { await context.owner.dispose(); }
  }
});

test('crash_before_and_after_worker_apply_gives_same_recovery', async () => {
  const bytes: Uint8Array[] = [];
  for (const before of [true, false]) {
    const context = await setup();
    const replay = spyOn(context.peerDoc, 'replayJson');
    context.controls.holdRequests = before;
    context.controls.holdReplies = !before;
    try {
      context.access.setSlideNotes(context.access.snapshot().slides[0].id, 'uncertain worker');
      await context.observed;
      context.crash();
      await expect(context.owner.flush()).rejects.toMatchObject({ code: 'editor-failed' });
      bytes.push((await context.owner.recoverySave()).bytes);
      await recovery(context, 'uncertain worker');
      expect(replay).toHaveBeenCalledTimes(1);
    } finally { replay.mockRestore(); await context.owner.dispose(); }
  }
  expect(bytes[0]).toEqual(bytes[1]);
});

test('terminal_failure_rejects_all_waiters', async () => {
  const context = await setup();
  context.controls.holdReplies = true;
  try {
    context.access.setSlideNotes(context.access.snapshot().slides[0].id, 'last accepted');
    const waiters = [context.owner.flush(), context.owner.saveAsync(), context.owner.flush()];
    const settled = Promise.allSettled(waiters);
    await context.observed;
    context.crash();
    expect((await settled).map((result) => result.status)).toEqual(['rejected', 'rejected', 'rejected']);
    expect(context.errors).toBe(1);
    expect(() => context.access.undo()).toThrow('failed');
    await recovery(context, 'last accepted');
  } finally { await context.owner.dispose(); }
});

test('recovery_serialization_failure_is_retryable', async () => {
  const context = await setup();
  context.access.setSlideNotes(context.access.snapshot().slides[0].id, 'retry recovery');
  context.crash();
  const save = spyOn(context.peerDoc, 'saveBytes').mockImplementationOnce(() => { throw new Error('serialization'); });
  const free = spyOn(context.peerDoc, 'free');
  try {
    await expect(context.owner.recoverySave()).rejects.toThrow('serialization');
    expect(context.owner.state.stage).toBe('failed');
    expect(free).not.toHaveBeenCalled();
    save.mockImplementationOnce(() => { throw new Error('retirement serialization'); });
    await context.owner.dispose();
    expect(free).not.toHaveBeenCalled();
    await recovery(context, 'retry recovery');
    expect(free).toHaveBeenCalledTimes(1);
  } finally { save.mockRestore(); free.mockRestore(); await context.owner.dispose(); }
});

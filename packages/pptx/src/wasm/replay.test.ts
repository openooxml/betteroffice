import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import {
  adoptPresentationPeerIdentity,
  initWasm,
  openPresentation,
  openPresentationPeerDeck,
  openPresentationReplayBaseline,
  presentationPeerDisplayListJson,
  presentationPeerHydration,
  presentationPeerMetadata,
  PresentationPeerError,
  registerPresentationPeerFonts,
  replayPresentation,
} from './loader';
import type {
  OpenPresentationOptions,
  PresentationHandle,
  PresentationPeerHandle,
  PresentationReplayEnvelope,
  PresentationReplayOp,
  PresentationReplayReply,
} from './loader';

const root = resolve(import.meta.dir, '../../../..');
const handles = new Set<PresentationPeerHandle>();
let source: Uint8Array;
let font: Uint8Array;

beforeAll(async () => {
  const [wasm, deck, bytes] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  source = deck;
  font = bytes;
});

afterEach(() => {
  for (const handle of handles) handle.dispose();
  handles.clear();
});

function keep<T extends PresentationPeerHandle>(handle: T): T {
  handles.add(handle);
  return handle;
}

function fonts(): OpenPresentationOptions {
  return {
    fonts: [{ family: 'Liberation Sans', bytes: font }],
    fallbackFonts: [{ family: 'Liberation Sans', bytes: font }],
  };
}

async function baseline(options: OpenPresentationOptions = {}) {
  const worker = keep(openPresentationReplayBaseline(source, { clientId: 5801, ...options }));
  await registerPresentationPeerFonts(worker);
  const identity = presentationPeerHydration(worker);
  return { worker, identity };
}

async function pair(options: OpenPresentationOptions = {}) {
  const { worker, identity } = await baseline(options);
  const peer = keep(openPresentationPeerDeck(source, identity, options));
  await registerPresentationPeerFonts(peer);
  adoptPresentationPeerIdentity(peer);
  let sequence = 0;
  function equal(prefix = 'baseline') {
    try {
      expect(peer.clientId).toBe(worker.clientId);
      expect(peer.version()).toBe(worker.version());
      expect(peer.snapshot()).toEqual(worker.snapshot());
      expect(peer.readContent()).toEqual(worker.readContent());
      expect(peer.listProposals()).toEqual(worker.listProposals());
      expect(peer.canUndo()).toBe(worker.canUndo());
      expect(peer.canRedo()).toBe(worker.canRedo());
      expect(peer.undoCaptureMode()).toBe('manual');
      expect(worker.undoCaptureMode()).toBe('manual');
      expect(peer.save()).toEqual(worker.save());
      expect(presentationPeerMetadata(peer)).toEqual(presentationPeerMetadata(worker));
      for (const slide of peer.snapshot().slides) {
        for (const shape of slide.shapes) {
          for (const story of shape.textStories) {
            expect(peer.story(story.id)).toEqual(worker.story(story.id));
          }
        }
      }
    } catch (error) {
      if (error instanceof Error) error.message = `Replay parity at ${prefix}: ${error.message}`;
      throw error;
    }
  }
  function run(op: PresentationReplayOp) {
    const envelope = { sequence: sequence + 1, baseVersion: peer.version(), op };
    const reply = replayPresentation(peer, envelope);
    const acknowledged = replayPresentation(worker, { ...envelope, expectedOutcome: reply.outcome });
    expect(acknowledged).toEqual(reply);
    expect(reply.version).toBe(peer.version());
    expect(reply.engineVersion).toBe(peer.version());
    if (reply.consumed) sequence += 1;
    equal(`${envelope.sequence} (${op.method})`);
    return reply;
  }
  equal();
  return { worker, peer, identity, run, equal };
}

function refusal(operation: () => unknown, code: string) {
  let caught: unknown;
  try { operation(); } catch (error) { caught = error; }
  expect(caught).toBeInstanceOf(PresentationPeerError);
  expect((caught as PresentationPeerError).name).toBe('PresentationPeerError');
  expect((caught as PresentationPeerError).code).toBe(code);
  expect((caught as PresentationPeerError).message.length).toBeGreaterThan(0);
}

test('all_supported_ops_match_across_handles', async () => {
  const { worker, peer, run } = await pair(fonts());
  const seen = new Set<string>();
  const apply = (op: PresentationReplayOp) => {
    seen.add(op.method);
    run({ method: 'addUndoBoundary', args: [] });
    return run(op);
  };
  apply({ method: 'insertSlide', args: [1] });
  const slide = peer.snapshot().slides[1].id;
  apply({ method: 'moveSlide', args: [slide, 0] });
  apply({ method: 'setSlideNotes', args: [slide, 'Replay notes'] });
  const rect = { x: 100000, y: 100000, width: 2000000, height: 900000 };
  apply({ method: 'addTextBox', args: [slide, { name: 'Replay text', rect, text: 'Hello', style: {} }] });
  const textbox = peer.snapshot().slides[0].shapes.find((shape) => shape.name === 'Replay text')!;
  const story = textbox.textStories[0].id;
  const peerAnchor = peer.anchorCaret(story, 2);
  const workerAnchor = worker.anchorCaret(story, 2);
  apply({ method: 'insertText', args: [story, 0, 'Hi ', { bold: true }] });
  expect(peer.resolveCaretAnchor(peerAnchor)).toBe(worker.resolveCaretAnchor(workerAnchor));
  apply({ method: 'formatText', args: [story, 0, 2, { italic: true, color: '#112233', fontSizePt: 18 }] });
  apply({ method: 'deleteText', args: [story, 0, 1] });
  apply({ method: 'insertParagraphBreak', args: [story, 2] });
  apply({ method: 'setParagraphAlignment', args: [story, 0, 2, 'ctr'] });
  apply({ method: 'addShape', args: [slide, { name: 'Replay shape', geometry: 'roundRect', rect, fill: '#123456' }] });
  const shape = peer.snapshot().slides[0].shapes.find((shape) => shape.name === 'Replay shape')!.id;
  apply({ method: 'moveShape', args: [slide, shape, 200000, 300000] });
  apply({ method: 'resizeShape', args: [slide, shape, 2100000, 1000000] });
  apply({ method: 'setShapeRect', args: [slide, shape, rect] });
  apply({ method: 'setShapeFill', args: [slide, shape, '#654321'] });
  apply({ method: 'setShapeStroke', args: [slide, shape, { color: '#abcdef', widthPt: 2 }] });
  apply({ method: 'setShapeAdjust', args: [slide, shape, { adj: 25000 }] });
  apply({ method: 'bringShapeToFront', args: [slide, textbox.id] });
  apply({ method: 'sendShapeToBack', args: [slide, textbox.id] });
  apply({ method: 'bringShapeForward', args: [slide, textbox.id] });
  apply({ method: 'sendShapeBackward', args: [slide, textbox.id] });
  apply({ method: 'addPicture', args: [slide, {
    name: 'Replay picture', rect, contentType: 'image/png',
    mediaBase64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=',
  }] });
  apply({ method: 'setCommentFlavor', args: ['modern'] });
  apply({ method: 'addComment', args: [slide, {
    author: 'Replay', initials: 'RP', text: 'Root', created: '2026-10-05T12:00:00Z', xEmu: 2, yEmu: 3,
  }] });
  const comment = peer.snapshot().comments!.find((item) => item.text === 'Root')!.id;
  apply({ method: 'replyToComment', args: [comment, { author: 'Reply', text: 'Child', created: '2026-10-05T12:01:00Z' }] });
  apply({ method: 'setCommentStatus', args: [comment, true] });
  apply({ method: 'setCommentPosition', args: [comment, { xEmu: 100, yEmu: 200 }] });
  apply({ method: 'removeComment', args: [comment] });
  expect(peer.snapshot().comments ?? []).toEqual([]);
  apply({ method: 'propose', args: ['agent', 'Review', [{ type: 'setSlideNotes', slideId: slide, text: 'Accepted' }]] });
  apply({ method: 'acceptProposal', args: [peer.listProposals()[0].id] });
  apply({ method: 'propose', args: ['agent', null, [{ type: 'setSlideNotes', slideId: slide, text: 'Rejected' }]] });
  apply({ method: 'rejectProposal', args: [peer.listProposals()[0].id] });
  apply({ method: 'applyEdits', args: [{
    expectVersion: peer.version(), history: 'separate',
    steps: [{ op: 'setSlideNotes', target: { slideId: slide }, text: 'Batch notes' }],
  }] });
  apply({ method: 'removeShape', args: [slide, shape] });
  apply({ method: 'deleteSlide', args: [slide] });
  apply({ method: 'addUndoBoundary', args: [] });
  const finalSnapshot = peer.snapshot();
  const finalBytes = peer.save();
  let undoCount = 0;
  while (peer.canUndo()) {
    expect(undoCount++).toBeLessThan(100);
    expect(apply({ method: 'undo', args: [] }).outcome.applied).toBe(true);
  }
  let redoCount = 0;
  while (peer.canRedo()) {
    expect(redoCount++).toBeLessThan(100);
    expect(apply({ method: 'redo', args: [] }).outcome.applied).toBe(true);
  }
  expect(redoCount).toBe(undoCount);
  expect(peer.snapshot()).toEqual(finalSnapshot);
  expect(peer.save()).toEqual(finalBytes);
  expect([...seen].sort()).toEqual([
    'insertText', 'deleteText', 'formatText', 'insertParagraphBreak', 'setParagraphAlignment',
    'insertSlide', 'deleteSlide', 'moveSlide', 'setSlideNotes', 'addTextBox', 'addShape', 'addPicture',
    'removeShape', 'moveShape', 'resizeShape', 'setShapeRect', 'setShapeFill', 'setShapeStroke',
    'setShapeAdjust', 'bringShapeToFront', 'sendShapeToBack', 'bringShapeForward', 'sendShapeBackward',
    'addComment', 'replyToComment', 'setCommentStatus', 'setCommentPosition', 'removeComment',
    'setCommentFlavor', 'propose', 'acceptProposal', 'rejectProposal', 'applyEdits',
    'addUndoBoundary', 'undo', 'redo',
  ].sort());
}, 60_000);

test('invalid_creation_then_valid_creation_keeps_ids', async () => {
  const { peer, worker, run, equal } = await pair();
  const slide = peer.snapshot().slides[0].id;
  const rect = { x: 0, y: 0, width: 1000000, height: 1000000 };
  const invalid: PresentationReplayOp[] = [
    { method: 'insertSlide', args: [999999] },
    { method: 'addTextBox', args: ['missing', { name: 'Refused', rect, text: 'Text', style: {} }] },
    { method: 'addComment', args: ['missing', { author: 'A', text: 'Refused', created: '2026-10-05T12:00:00Z' }] },
    { method: 'replyToComment', args: ['missing', { author: 'A', text: 'Refused', created: '2026-10-05T12:00:00Z' }] },
  ];
  const before = peer.snapshot();
  const version = peer.version();
  for (const op of invalid) {
    expect(() => replayPresentation(peer, { sequence: 1, baseVersion: version, op })).toThrow(PresentationPeerError);
    expect(peer.version()).toBe(version);
    expect(peer.snapshot()).toEqual(before);
    expect(peer.canUndo()).toBe(false);
    expect(peer.canRedo()).toBe(false);
    equal();
  }
  run({ method: 'setCommentFlavor', args: ['modern'] });
  run({ method: 'insertSlide', args: [1] });
  run({ method: 'addTextBox', args: [slide, { name: 'Valid', rect, text: 'Text', style: {} }] });
  run({ method: 'addComment', args: [slide, { author: 'A', text: 'Valid', created: '2026-10-05T12:00:00Z' }] });
  run({ method: 'replyToComment', args: [peer.snapshot().comments![0].id, { author: 'A', text: 'Reply', created: '2026-10-05T12:01:00Z' }] });
  expect(peer.snapshot()).toEqual(worker.snapshot());
});

test('ordinary_open_keeps_session_scoped_versions', () => {
  const first = keep(openPresentation(source, { clientId: 5802 }));
  const second = keep(openPresentation(source, { clientId: 5802 }));
  expect(first.version()).not.toBe(second.version());
  expect(first.snapshot()).toEqual(second.snapshot());
  const prior = first.version();
  first.setSlideNotes(first.snapshot().slides[0].id, 'Ordinary edit');
  expect(first.version()).not.toBe(prior);
  expect(second.snapshot().slides[0].notes).not.toBe('Ordinary edit');
  refusal(() => presentationPeerHydration(first), 'stage');
});

test('layout_and_proposal_previews_do_not_change_replay_state', async () => {
  const { peer, worker, run, equal } = await pair(fonts());
  const slide = peer.snapshot().slides[0];
  const shape = slide.shapes.find((item) => item.textStories.length > 0)!;
  const story = shape.textStories[0].id;
  run({ method: 'propose', args: ['agent', null, [{ type: 'replaceText', storyId: story, start: 0, end: 0, text: 'Preview ' }]] });
  const id = peer.listProposals()[0].id;
  const before = peer.snapshot();
  const version = peer.version();
  const bytes = peer.save();
  const proposals = peer.listProposals();
  const anchor = peer.anchorCaret(story, 0);
  for (let index = 0; index < 3; index += 1) {
    expect(presentationPeerDisplayListJson(peer, 0)).toBe(presentationPeerDisplayListJson(worker, 0));
    expect(peer.layoutSlide(0)).toEqual(worker.layoutSlide(0));
    expect(peer.hitTest(30, 30)).toEqual(worker.hitTest(30, 30));
    expect(peer.previewProposal(id)).toEqual(worker.previewProposal(id));
    expect(peer.layoutProposalSlide(id, 0)).toEqual(worker.layoutProposalSlide(id, 0));
    expect(peer.layoutProposalDiffSlide(id, 0)).toEqual(worker.layoutProposalDiffSlide(id, 0));
  }
  expect(peer.snapshot()).toEqual(before);
  expect(peer.version()).toBe(version);
  expect(peer.save()).toEqual(bytes);
  expect(peer.listProposals()).toEqual(proposals);
  expect(peer.resolveCaretAnchor(anchor)).toBe(0);
  equal();
  run({ method: 'acceptProposal', args: [id] });
  run({ method: 'addTextBox', args: [slide.id, {
    name: 'After previews', text: 'Same allocator', style: {},
    rect: { x: 0, y: 0, width: 1000000, height: 1000000 },
  }] });
  run({ method: 'propose', args: ['agent', null, [{ type: 'setSlideNotes', slideId: slide.id, text: 'Next proposal' }]] });
});

test('staged_hydration_refusals_dispose_without_adoption', async () => {
  const { worker, identity } = await baseline(fonts());
  const decoded = JSON.parse(identity);
  refusal(() => openPresentationPeerDeck(source, JSON.stringify({ ...decoded, schemaVersion: 99 })), 'schema');
  refusal(() => openPresentationPeerDeck(new Uint8Array([1, 2, 3]), identity), 'sourceMismatch');
  refusal(() => openPresentationReplayBaseline(source, { clientId: 0 }), 'clientId');
  refusal(() => openPresentationPeerDeck(source, identity, { initialUpdate: new Uint8Array([255]) }), 'initialUpdate');

  const skipped = keep(openPresentationPeerDeck(source, identity, fonts()));
  refusal(() => adoptPresentationPeerIdentity(skipped), 'stage');
  expect(() => skipped.snapshot()).toThrow('disposed');

  const reversed = keep(openPresentationPeerDeck(source, identity, {
    fonts: [], fallbackFonts: [{ family: 'Liberation Sans', bytes: font }, { family: 'Liberation Sans', bytes: font }],
  }));
  await expect(registerPresentationPeerFonts(reversed)).rejects.toMatchObject({ name: 'PresentationPeerError', code: 'fontOrder' });
  expect(() => reversed.snapshot()).toThrow('disposed');
  refusal(() => adoptPresentationPeerIdentity(reversed), 'stage');

  const badFont = keep(openPresentationPeerDeck(source, identity, {
    fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array([1, 2, 3]) }],
  }));
  await expect(registerPresentationPeerFonts(badFont)).rejects.toBeInstanceOf(PresentationPeerError);
  expect(() => badFont.snapshot()).toThrow('disposed');

  const mismatch = keep(openPresentationPeerDeck(source, JSON.stringify({ ...decoded, epoch: String(BigInt(decoded.epoch) + 1n) }), fonts()));
  await registerPresentationPeerFonts(mismatch);
  refusal(() => adoptPresentationPeerIdentity(mismatch), 'baselineMismatch');
  expect(() => mismatch.version()).toThrow('disposed');

  const valid = keep(openPresentationPeerDeck(source, identity, fonts()));
  await registerPresentationPeerFonts(valid);
  adoptPresentationPeerIdentity(valid);
  expect(valid.version()).toBe(worker.version());
  expect(valid.save()).toEqual(worker.save());
});

test('initial_update_reconstructs_the_same_baseline_with_a_fresh_client', async () => {
  const seed = keep(openPresentation(source, { clientId: 5803 }));
  seed.setSlideNotes(seed.snapshot().slides[0].id, 'Initial update notes');
  const initialUpdate = seed.encodeStateAsUpdate();
  refusal(() => openPresentationReplayBaseline(source, { clientId: seed.clientId, initialUpdate }), 'clientIdReuse');
  const { peer, run } = await pair({ initialUpdate });
  expect(peer.snapshot()).toEqual(seed.snapshot());
  run({ method: 'insertSlide', args: [1] });
});

test('premature_capture_and_disposal_during_font_hashing_refuse_cleanly', async () => {
  const premature = keep(openPresentationReplayBaseline(source));
  refusal(() => presentationPeerHydration(premature), 'stage');
  expect(() => premature.snapshot()).toThrow('disposed');
  const { identity } = await baseline(fonts());
  const peer = keep(openPresentationPeerDeck(source, identity, fonts()));
  const registration = registerPresentationPeerFonts(peer);
  peer.dispose();
  await expect(registration).rejects.toMatchObject({ name: 'PresentationPeerError', code: 'stage' });
  refusal(() => adoptPresentationPeerIdentity(peer), 'stage');
});

test('owned_font_subviews_and_order_survive_caller_mutation', async () => {
  const storage = new Uint8Array(font.length + 10);
  storage.set(font, 5);
  const view = storage.subarray(5, 5 + font.length);
  const options = { fonts: [{ family: 'Liberation Sans', bytes: view }], fallbackFonts: [{ family: 'Liberation Sans', bytes: view }] };
  const worker = keep(openPresentationReplayBaseline(source, options));
  storage.fill(0);
  options.fonts[0].family = 'Changed';
  await registerPresentationPeerFonts(worker);
  const identity = presentationPeerHydration(worker);
  const manifest = JSON.parse(identity).fonts;
  const fingerprint = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', font.slice())), (byte) => byte.toString(16).padStart(2, '0')).join('');
  expect(manifest).toEqual([
    { family: 'Liberation Sans', bold: false, italic: false, fallback: false, fingerprint },
    { family: 'Liberation Sans', bold: false, italic: false, fallback: true, fingerprint },
  ]);
  const peer = keep(openPresentationPeerDeck(source, identity, fonts()));
  await registerPresentationPeerFonts(peer);
  adoptPresentationPeerIdentity(peer);
  expect(peer.layoutSlide(0)).toEqual(worker.layoutSlide(0));
  const buffer = Buffer.from(font);
  const buffered = keep(openPresentationReplayBaseline(source, {
    fonts: [{ family: 'Liberation Sans', bytes: buffer }],
  }));
  buffer.fill(0);
  await registerPresentationPeerFonts(buffered);
  expect(JSON.parse(presentationPeerHydration(buffered)).fonts[0].fingerprint).toBe(fingerprint);
});

test('outcome_is_captured_before_listeners_and_replay_cannot_reenter', async () => {
  const { peer, worker } = await pair();
  const slide = peer.snapshot().slides[0].id;
  const envelope: PresentationReplayEnvelope = {
    sequence: 1, baseVersion: peer.version(), op: { method: 'setSlideNotes', args: [slide, 'Owned notes'] },
  };
  let captured: PresentationReplayReply | undefined;
  let listenerFailure: unknown;
  const order: string[] = [];
  const off = peer.onUpdate(() => {
    order.push('listener');
    try {
      expect(captured).toBeDefined();
      expect(captured!.version).toBe(peer.version());
      expect(captured!.engineVersion).toBe(peer.version());
      expect(peer.snapshot().slides[0].notes).toBe('Owned notes');
      expect(Object.isFrozen(captured)).toBe(true);
      expect(Object.isFrozen(captured!.outcome)).toBe(true);
      refusal(() => replayPresentation(peer, {
        sequence: 2, baseVersion: peer.version(), op: { method: 'setSlideNotes', args: [slide, 'Nested'] },
      }), 'stage');
    } catch (error) { listenerFailure = error; }
  });
  const original = structuredClone(envelope);
  const reply = replayPresentation(peer, envelope, (result) => {
    order.push('captured');
    captured = result;
    envelope.op.args[1] = 'Caller changed';
    refusal(() => replayPresentation(peer, original), 'stage');
  });
  expect(order).toEqual(['captured', 'listener']);
  expect(listenerFailure).toBeUndefined();
  off();
  expect(captured).toBe(reply);
  expect(replayPresentation(worker, { ...original, expectedOutcome: reply.outcome })).toEqual(reply);
  expect(peer.save()).toEqual(worker.save());
  expect(replayPresentation(peer, {
    sequence: 2, baseVersion: peer.version(), op: { method: 'setSlideNotes', args: [slide, 'Next'] },
  }).sequence).toBe(2);
});

test('ordinary_mutations_are_refused_on_retained_handles_at_every_stage', async () => {
  const { worker, identity } = await baseline(fonts());
  const peer = keep(openPresentationPeerDeck(source, identity, fonts()));
  const methods = [
    'insertText', 'deleteText', 'formatText', 'insertParagraphBreak', 'setParagraphAlignment',
    'insertSlide', 'deleteSlide', 'moveSlide', 'setSlideNotes', 'addTextBox', 'addShape', 'addPicture',
    'removeShape', 'moveShape', 'resizeShape', 'setShapeRect', 'setShapeFill', 'setShapeStroke',
    'setShapeAdjust', 'bringShapeToFront', 'sendShapeToBack', 'bringShapeForward', 'sendShapeBackward',
    'addComment', 'replyToComment', 'setCommentStatus', 'setCommentPosition', 'removeComment',
    'setCommentFlavor', 'propose', 'acceptProposal', 'rejectProposal', 'applyEdits', 'addUndoBoundary',
    'undo', 'redo', 'insertTextProfiled', 'deleteTextProfiled', 'insertSlideProfiled',
    'addTextBoxProfiled', 'moveShapeProfiled', 'undoProfiled', 'applyUpdate', 'setUndoCaptureMode',
    'registerFont', 'registerFallbackFont',
  ];
  const blocked = (handle: PresentationPeerHandle) => {
    for (const method of methods) {
      refusal(() => (handle as unknown as Record<string, () => unknown>)[method](), 'stage');
    }
  };
  blocked(worker);
  blocked(peer);
  refusal(() => replayPresentation(peer, { sequence: 1, baseVersion: peer.version(), op: { method: 'undo', args: [] } }), 'stage');
  const registration = registerPresentationPeerFonts(peer);
  blocked(peer);
  await registration;
  blocked(peer);
  adoptPresentationPeerIdentity(peer);
  blocked(peer);
  const retained = peer as PresentationHandle;
  expect(retained.snapshot()).toEqual(worker.snapshot());
  blocked(retained);
  peer.dispose();
  blocked(retained);
});

test('replay_refusals_preserve_codes_and_noops_consume_sequence', async () => {
  const { peer, worker, identity, run, equal } = await pair();
  const baseVersion = peer.version();
  const op: PresentationReplayOp = { method: 'undo', args: [] };
  refusal(() => replayPresentation(peer, { sequence: 2, baseVersion, op }), 'sequence');
  refusal(() => replayPresentation(peer, { sequence: 1, baseVersion: 'stale', op }), 'version');
  refusal(() => replayPresentation(peer, {
    sequence: 1, baseVersion, op: { method: 'deleteText', args: ['missing', -1, 0] },
  }), 'invalidJson');
  refusal(() => replayPresentation(peer, {
    sequence: 1, baseVersion, op: { method: 'setShapeRect', args: ['missing', 'missing', { x: 0.5, y: 0, width: 1, height: 1 }] },
  }), 'arguments');
  refusal(() => replayPresentation(peer, {
    sequence: 1, baseVersion, op: { method: 'unknown', args: [] } as unknown as PresentationReplayOp,
  }), 'method');
  equal();
  for (const noOp of [op, { method: 'redo', args: [] }, { method: 'rejectProposal', args: ['missing'] }] as PresentationReplayOp[]) {
    const reply = run(noOp);
    expect(reply.consumed).toBe(true);
    expect(reply.outcome.applied).toBe(false);
    expect(reply.version).toBe(baseVersion);
  }
  const reply = run({ method: 'addUndoBoundary', args: [] });
  expect(reply.sequence).toBe(4);
  expect(reply.revision).toBe(4);
  expect(presentationPeerHydration(worker)).toBe(identity);
});

import { expect, mock, spyOn, test } from 'bun:test';
import { CommandOutcomeUnknown } from '@betteroffice/pptx';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { encodeAwarenessUpdate, encodeUpdate } from '../../../pptx/src/collaboration/protocol';
import { createPresentationCommandProjection,
  type PresentationSharedEnvelope } from '../../../pptx/src/session/commandSharedAdapter';
import type {
  PptxReadRequest,
  PresentationCommandAck,
  PresentationCommandBasis,
  PresentationCommandClient,
  PresentationCommandFrame,
  PresentationCommandFailure,
  PresentationCommandIntent,
  PresentationCommandEnvelope,
  PresentationCommandOpenResult,
  PresentationCommandReadResult,
  PresentationCommandSelection,
  PresentationCommandThumbnailResult,
  SlideDisplayList, TextBoxPrimitive,
} from '@betteroffice/pptx';
import {
  createPeerlessViewStore,
  installCoveringFrame,
  hitConfirmedGeometry,
  PEERLESS_ACTIVE_CONTEXT_BYTES,
  PEERLESS_PENDING_INTENTS,
  peerlessInputReference,
  peerlessKnownState,
  type PeerlessConfirmedFrame,
  type PeerlessViewSnapshot,
} from './peerlessViewStore';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function selection(position = 1): PresentationCommandSelection {
  return { slideId: 'slide', shapeId: 'box', storyId: 'story',
    anchor: position, focus: position, affinity: 'downstream' };
}

function frame(
  text = 'abc', revision = 0, coversThroughSeq = 0,
  overrides: Partial<PresentationCommandFrame> = {},
): PresentationCommandFrame {
  const box: TextBoxPrimitive = {
    kind: 'textBox', objectId: 1, shapeId: 'box', storyId: 'story',
    x: 20, y: 40, w: 300, h: 40, anchor: 'top',
    paragraphs: [{ level: 0, runs: [{ text, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
    lines: [{ x: 20, y: 40, width: text.length * 10, height: 20, baseline: 55,
      start: 0, end: text.length,
      caretStops: Array.from({ length: text.length + 1 }, (_, position) => ({ position, x: 20 + position * 10 })),
      runs: [{ text, start: 0, end: text.length, x: 20, width: text.length * 10,
        fontId: 0, fontFamily: 'Arial', fontSizePx: 16, bold: false, italic: false,
        underline: false, color: '#000000', glyphs: [] }] }],
  };
  const displayList: SlideDisplayList = {
    contractVersion: 1, width: 960, height: 540,
    background: { kind: 'solid', color: '#ffffff' }, primitives: [box, {
      ...box, objectId: 2, shapeId: 'other-box', storyId: 'other-story', x: 400,
      lines: box.lines.map((line) => ({ ...line, x: 400,
        caretStops: line.caretStops.map((stop) => ({ ...stop, x: stop.x + 380 })),
        runs: line.runs.map((run) => ({ ...run, x: 400 })),
      })),
    }, {
      ...box, objectId: 3, shapeId: 'coverage', storyId: 'coverage-story',
      lines: box.lines.map((line) => ({ ...line, start: 0, end: 17, width: 170,
        caretStops: Array.from({ length: 18 }, (_, position) => ({ position, x: 20 + position * 10 })),
        runs: [{ ...line.runs[0]!, text: 'abcXYZxy!?LASTabc', start: 0, end: 17, width: 170 }],
      })),
    }],
  };
  return { slideId: 'slide', slideIndex: 0, slideRevision: revision, coversThroughSeq,
    documentEpoch: 1, incarnation: 'worker', epoch: revision + 1, version: `v${revision}`,
    displayList: new TextEncoder().encode(JSON.stringify(displayList)).buffer,
    media: [], ...overrides };
}

function initial(text = 'abc'): PresentationCommandOpenResult & { frame: PresentationCommandFrame } {
  return { session: 'session', incarnation: 'worker', epoch: 1, throughSeq: 0, version: 'v0',
    selection: selection(), editState: { readOnly: false, canUndo: false, canRedo: false,
      toolbar: { target: selection(), storedMarks: {}, formatting: null, alignment: null,
        hyperlink: null, table: null, canEditText: true, ...{ echoRun: 'after' as const } } },
    outcomes: [], dirtySlideIds: [], frame: frame(text), activeSlideId: 'slide',
    size: { width: 960, height: 540 }, slides: [
      { id: 'slide', index: 0, name: null, layoutPartPath: null },
      { id: 'other', index: 1, name: null, layoutPartPath: null },
    ] };
}

class FakeCommandClient implements PresentationCommandClient {
  private current: ReturnType<typeof initial>;
  private failure: PresentationCommandFailure | undefined;
  readonly sent: PresentationCommandIntent[] = [];
  readonly envelopes: PresentationCommandEnvelope[] = [];
  readonly replies = new Map<number, ReturnType<typeof deferred<PresentationCommandAck>>>();
  readonly ackListeners = new Set<(ack: PresentationCommandAck) => void>();
  readonly frameListeners = new Set<(frame: PresentationCommandFrame) => void>();
  readonly failureListeners = new Set<(error: PresentationCommandFailure) => void>();

  constructor(text = 'abc') { this.current = initial(text); }

  get state() { return this.current; }

  dispatch = mock((intent: PresentationCommandIntent, basis?: PresentationCommandBasis): Promise<PresentationCommandAck> => {
    if (this.failure) return Promise.reject(this.failure);
    this.sent.push(intent);
    this.envelopes.push({ ...intent, session: this.state.session, incarnation: this.state.incarnation,
      seq: this.sent.length, basis: basis ?? { epoch: this.state.epoch, version: this.state.version },
      ...('at' in intent.payload && intent.payload.at?.source === 'afterSeq' ? { afterSeq: intent.payload.at.seq } : {}) });
    const reply = deferred<PresentationCommandAck>();
    this.replies.set(this.sent.length, reply);
    return reply.promise;
  });
  flush = mock((_seq?: number): Promise<PresentationCommandAck> => Promise.resolve(this.state));
  save = mock((): Promise<Uint8Array> => Promise.resolve(new Uint8Array()));
  read = mock((_request?: PptxReadRequest): Promise<PresentationCommandReadResult> =>
    Promise.reject(new Error('main-thread story read')));
  thumbnail = mock((slideId: string, knownRevision?: number): Promise<PresentationCommandThumbnailResult> =>
    Promise.resolve({ session: 'session', incarnation: 'worker', epoch: 1,
      slideId, slideRevision: knownRevision ?? 0, status: 'unchanged' }));
  dispose = mock((): Promise<void> => Promise.resolve());

  subscribe(listener: (ack: PresentationCommandAck) => void) {
    this.ackListeners.add(listener);
    return () => { this.ackListeners.delete(listener); };
  }

  subscribeFrame(listener: (frame: PresentationCommandFrame) => void) {
    this.frameListeners.add(listener);
    return () => { this.frameListeners.delete(listener); };
  }

  subscribeMedia(_listener: (media: PresentationCommandFrame['media']) => void) { return () => {}; }

  subscribeFailure(listener: (error: PresentationCommandFailure) => void) {
    if (this.failure) listener(this.failure);
    this.failureListeners.add(listener);
    return () => { this.failureListeners.delete(listener); };
  }

  fail(error: PresentationCommandFailure) {
    this.failure = error;
    for (const listener of this.failureListeners) listener(error);
    for (const reply of this.replies.values()) reply.reject(error);
    this.replies.clear();
  }

  ack(throughSeq: number, overrides: Partial<PresentationCommandAck> = {}) {
    const { frame: _frame, slides: _slides, size: _size, ...base } = this.state;
    const ack: PresentationCommandAck = { ...base, throughSeq, version: `v${throughSeq}`,
      selection: selection(1 + throughSeq), dirtySlideIds: ['slide'],
      outcomes: Array.from({ length: throughSeq }, (_, index): PresentationCommandAck['outcomes'][number] => ({ seq: index + 1,
        status: 'applied', committed: true })), ...overrides };
    const currentIdentity = ack.session === this.state.session && ack.incarnation === this.state.incarnation
      && ack.epoch === this.state.epoch;
    if (currentIdentity && throughSeq >= this.state.throughSeq) {
      this.current = { ...this.current, ...ack, frame: ack.frame ?? this.current.frame };
    }
    for (const listener of this.ackListeners) listener(ack);
    if (ack.frame) this.deliverFrame(ack.frame);
    if (currentIdentity) {
      for (const [seq, reply] of this.replies) {
        if (seq <= throughSeq) { reply.resolve(ack); this.replies.delete(seq); }
      }
    }
    return ack;
  }

  deliverFrame(value: PresentationCommandFrame) {
    if (value.incarnation === this.current.incarnation && value.documentEpoch === this.current.epoch
      && value.slideId === this.current.activeSlideId && value.slideRevision > this.current.frame.slideRevision) {
      this.current = { ...this.current, frame: value };
    }
    for (const listener of this.frameListeners) listener(value);
  }
}

test('refused local input surfaces through onError exactly once', async () => {
  const client = new FakeCommandClient();
  const onError = mock((_error: Error) => {});
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} }, { onError });
  try {
    const input = store.dispatch(insert('NEXT'));
    const ack = client.ack(1, { selection: selection(), dirtySlideIds: [], outcomes: [
      { seq: 1, status: 'refused', reason: 'stale-basis', committed: false },
    ] });
    await input;
    client.ack(1, ack);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0][0].message).toContain('stale-basis');
  } finally { store.dispose(); }
});

test('refusals returned to the dispatching caller stay off onError', async () => {
  const client = new FakeCommandClient();
  const onError = mock((_error: Error) => {});
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} }, { onError });
  try {
    const input = store.dispatch({ kind: 'invoke', payload: { method: 'propose', args: ['Writer', null, []],
      inputTime: 100, undoGroup: 'object' } }, { reportRefusal: false });
    client.ack(1, { selection: selection(), dirtySlideIds: [], outcomes: [
      { seq: 1, status: 'refused', reason: 'invalid-payload', committed: false },
    ] });
    expect((await input).outcomes).toEqual([{ seq: 1, status: 'refused', reason: 'invalid-payload', committed: false }]);
    expect(onError).not.toHaveBeenCalled();
    expect(store.getSnapshot().inputFailure).toBeUndefined();
    expect(store.getSnapshot().editState.readOnly).toBe(false);
  } finally { store.dispose(); }
});

function insert(text: string, afterSeq?: number, position = 1): PresentationCommandIntent {
  return { kind: 'insertText', payload: { text, inputTime: 100, undoGroup: 'typing',
    at: afterSeq === undefined ? { source: 'confirmed', selection: selection(position) }
      : { source: 'afterSeq', seq: afterSeq } } };
}

function setup(text = 'abc', echoEnabled = true) {
  const client = new FakeCommandClient(text);
  const paints: { kind: string; view: PeerlessViewSnapshot }[] = [];
  const paintFrame = mock((_frame: PeerlessConfirmedFrame, view: PeerlessViewSnapshot) => {
    paints.push({ kind: 'frame', view });
  });
  const paintEcho = mock((view: PeerlessViewSnapshot) => { paints.push({ kind: 'echo', view }); });
  const restoreLine = mock((view: PeerlessViewSnapshot) => { paints.push({ kind: 'restore', view }); });
  const store = createPeerlessViewStore(client, { paintFrame, paintEcho, restoreLine }, { echoEnabled });
  return { client, store, paints, paintFrame, paintEcho, restoreLine };
}

for (const kind of ['cursor', 'awareness', 'noop-update', 'update'] as const) test(`pipe ${kind} preserves a local caret until the document changes`, async () => {
  const { client, store } = setup('abc', false);
  try {
    store.setSelection(selection(3));
    const ops: Extract<PresentationCommandIntent, { kind: 'collabPipe' }>['payload']['ops'] = kind === 'cursor'
      ? [{ op: 'cursor', seq: 1, cursor: { slideId: 'slide' }, broadcast: true }]
      : [{ op: 'event', seq: 1, event: { type: 'message', data: new Uint8Array(kind === 'awareness'
        ? encodeAwarenessUpdate([{ clientId: 9, clock: 1, state: null }])
        : encodeUpdate(Uint8Array.of(0, 0))).buffer } }];
    const operation = store.dispatch({ kind: 'collabPipe', payload: { provider: 1, ops,
      inputTime: 1, undoGroup: 'remote' } });
    expect(store.getSnapshot().selection).toEqual(selection(3));
    if (kind === 'cursor') expect(peerlessInputReference(store)).toEqual({ source: 'confirmed', selection: selection(3) });
    const changed = kind === 'update';
    client.ack(1, { selection: selection(1), dirtySlideIds: changed ? ['slide'] : [],
      ...(changed ? { historyThroughSeq: 1 } : {}),
      outcomes: [{ seq: 1, status: changed ? 'applied' : 'noop', committed: changed }],
      frame: frame(changed ? 'Xabc' : 'abc', changed ? 1 : 0, 1) });
    await operation;
    expect(store.getSnapshot().selection).toEqual(selection(changed ? 1 : 3));
    expect(peerlessInputReference(store)).toEqual({ source: 'confirmed', selection: selection(changed ? 1 : 3) });
  } finally { store.dispose(); }
});

for (const queued of [false, true]) for (const kind of ['cursor', 'message'] as const) {
  test(`pipe_${kind}_${queued ? 'queued' : 'sent'}_does_not_publish`, async () => {
    const { client, store, paints } = setup('abc', false);
    try {
      store.setSelection(selection(3));
      const predecessor = queued ? store.dispatch(insert('X')) : undefined;
      const publications = mock(() => {});
      store.subscribePresentation(publications);
      const before = store.getSnapshot();
      const paintCount = paints.length;
      const ops: Extract<PresentationCommandIntent, { kind: 'collabPipe' }>['payload']['ops'] = kind === 'cursor'
        ? [{ op: 'cursor', seq: 1, cursor: { slideId: 'slide' }, broadcast: true }]
        : [{ op: 'event', seq: 1, event: { type: 'message', data: Uint8Array.of(255).buffer } }];
      const operation = store.dispatch({ kind: 'collabPipe', payload: { provider: 1, ops,
        inputTime: 1, undoGroup: 'remote' } });
      expect(publications).not.toHaveBeenCalled();
      expect(paints).toHaveLength(paintCount);
      expect(store.getSnapshot().selection).toBe(before.selection);
      expect(store.getSnapshot().caret).toBe(before.caret);
      if (queued) {
        expect(client.sent).toHaveLength(1);
        expect(store.getSnapshot().waitingInputCount).toBe(1);
        client.ack(1, { frame: frame('aXbc', 1, 1) });
      }
      const seq = queued ? 2 : 1;
      client.ack(seq, { version: queued ? 'v1' : 'v0', dirtySlideIds: [],
        outcomes: [{ seq, status: 'noop', committed: false }],
        frame: frame(queued ? 'aXbc' : 'abc', queued ? 1 : 0, seq, { epoch: seq + 1 }) });
      await Promise.all([predecessor, operation]);
    } finally { store.dispose(); }
  });
}

test('pipe_opaque_message_keeps_typing_order_after_noop_ack', async () => {
  const { client, store } = setup('abc', false);
  try {
    const pipe = store.dispatch({ kind: 'collabPipe', payload: { provider: 1,
      ops: [{ op: 'event', seq: 1, event: { type: 'message', data: Uint8Array.of(255).buffer } }],
      inputTime: 1, undoGroup: 'remote' } });
    expect(peerlessKnownState(store)).toBe(false);
    expect(peerlessInputReference(store)).toEqual({ source: 'afterSeq', seq: 1 });
    const typed = store.dispatch(insert('X', 1));
    expect(client.sent).toHaveLength(1);
    client.ack(1, { version: 'v0', selection: selection(), dirtySlideIds: [],
      outcomes: [{ seq: 1, status: 'noop', committed: false }], frame: frame('abc', 0, 1, { epoch: 2 }) });
    expect(client.envelopes[1]).toMatchObject({ kind: 'insertText', afterSeq: 1,
      payload: { at: { source: 'afterSeq', seq: 1 } } });
    client.ack(2, { version: 'v1', frame: frame('aXbc', 1, 2) });
    await Promise.all([pipe, typed]);
    expect(peerlessKnownState(store)).toBe(true);
    expect(store.getSnapshot().pending).toEqual([]);
  } finally { store.dispose(); }
});

test('activation_queued_behind_300_mutations_uses_send_time_basis', async () => {
  const { client, store } = setup('abc', false);
  const operations: Promise<PresentationCommandAck>[] = [];
  try {
    for (let index = 0; index < 300; index += 1) {
      const operation = store.dispatch({ kind: 'insertText', payload: { text: 'x', inputTime: 100,
        undoGroup: `edit-${index}`, at: { source: 'confirmed', selection: selection() } } });
      void operation.catch(() => {});
      operations.push(operation);
    }
    const navigation = store.dispatch({ kind: 'activateSlide', payload: { slideId: 'other' } });
    void navigation.catch(() => {});
    operations.push(navigation);
    expect(client.sent).toHaveLength(1);
    for (let seq = 1; seq <= 300; seq += 1) {
      client.ack(seq, { frame: frame('abc', seq, seq) });
    }
    expect(client.envelopes[client.envelopes.length - 1]).toMatchObject({ kind: 'activateSlide', basis: { epoch: 1, version: 'v300' } });
    client.ack(301, { activeSlideId: 'other', selection: null,
      frame: frame('other', 0, 301, { slideId: 'other', slideIndex: 1, version: 'v300' }) });
    expect(store.getSnapshot().confirmed?.frame.slideId).toBe('other');
  } finally {
    store.dispose();
    await Promise.allSettled(operations);
  }
});

test('activation_clears_composition_and_preedit', async () => {
  const { client, store } = setup();
  try {
    store.setFocused(true);
    store.beginComposition('activation-ime', 100);
    store.updateComposition('xy');
    client.ack(1, { selection: selection(), outcomes: [{ seq: 1, status: 'applied', committed: false }],
      frame: frame('abc', 0, 1, { epoch: 2, version: 'v1' }) });
    expect(store.getSnapshot().composition).toMatchObject({ id: 'activation-ime', text: 'xy' });
    expect(store.getSnapshot().echo?.text).toBe('axybc');
    const operation = store.dispatch({ kind: 'activateSlide', payload: { slideId: 'other' } });
    void operation.catch(() => {});
    client.ack(2, { activeSlideId: 'other', selection: null, dirtySlideIds: [],
      outcomes: [{ seq: 2, status: 'applied', committed: false }],
      frame: frame('other', 0, 2, { slideId: 'other', slideIndex: 1, epoch: 3, version: 'v2' }) });
    expect(store.getSnapshot().confirmed?.frame.slideId).toBe('other');
    expect(store.getSnapshot().selection).toBeNull();
    expect(store.getSnapshot().composition).toBeNull();
    expect(store.getSnapshot().echo).toBeNull();
  } finally { store.dispose(); }
});

for (const [mode, createClient] of [['command', createPresentationCommandClient],
  ['shared', createSharedPresentationCommandClient]] as const) {
  test(`pipe_noop_ack_reuses_installed_content_without_paint_${mode}`, async () => {
    const state = initial();
    const projection = createPresentationCommandProjection();
    projection.initialize(state.version);
    let deliver: (message: unknown) => void = () => {};
    const held: (() => void)[] = [];
    const transport: SessionTransport = {
      post(message) {
        if (!isClientMessage(message) || message.kind !== 'call') return;
        const reply = (value: unknown) => deliver({ protocol: 1, kind: 'reply', id: message.id, ok: true, value });
        if (message.method === 'openCommandEditor') reply({ ...state, shared: projection.ack(state) });
        else if (message.method === 'command' || message.method === 'edit') {
          const command = message.args[0] as PresentationCommandEnvelope;
          const ack: PresentationCommandAck = { ...state, throughSeq: command.seq,
            historyThroughSeq: 0, frame: frame('abc', 0, command.seq, { epoch: command.seq + 1 }),
            outcomes: [{ seq: command.seq, status: 'noop', committed: false }] };
          held.push(() => reply(mode === 'shared' ? projection.ack(ack) : ack));
        }
      },
      listen(listener) { deliver = listener; return () => {}; },
      onError() { return () => {}; },
      close() {},
    };
    const client = await createClient(new Uint8Array(), { session: state.session, incarnation: state.incarnation }, transport);
    const paintFrame = mock((_frame: PeerlessConfirmedFrame, _view: PeerlessViewSnapshot) => {});
    const paintEcho = mock((_view: PeerlessViewSnapshot) => {});
    const restoreLine = mock((_view: PeerlessViewSnapshot) => {});
    const store = createPeerlessViewStore(client, { paintFrame, paintEcho, restoreLine });
    store.setSelection(selection(3));
    const before = store.getSnapshot();
    const publications = mock(() => {});
    store.subscribePresentation(publications);
    const parse = spyOn(JSON, 'parse');
    try {
      for (let seq = 1; seq <= 2; seq++) {
        const operation = store.dispatch({ kind: 'collabPipe', payload: { provider: 1,
          ops: [{ op: 'event', seq, event: { type: 'message', data: new Uint8Array(encodeUpdate(Uint8Array.of(0, 0))).buffer } }],
          inputTime: seq, undoGroup: 'remote' } });
        for (let turn = 0; turn < 32; turn++) await Promise.resolve();
        expect(held).toHaveLength(seq);
        expect(peerlessInputReference(store)).toEqual({ source: 'afterSeq', seq });
        held[seq - 1]!();
        await operation;
        await store.fence();
        expect(store.getSnapshot().confirmed).toBe(before.confirmed);
        expect(store.getSnapshot().confirmed?.frame).toMatchObject({ coversThroughSeq: seq, epoch: seq + 1 });
        expect(store.getSnapshot().confirmed?.frame.displayList).toBe(state.frame.displayList);
        expect(store.getSnapshot().activeLine).toBe(before.activeLine);
        expect(store.getSnapshot().caret).toBe(before.caret);
        expect(store.getSnapshot().selection).toEqual(selection(3));
        expect(store.getSnapshot().presentedThroughSeq).toBe(seq);
        expect(store.getSnapshot().appliedThroughSeq).toBe(seq);
        expect(store.getSnapshot().pending).toEqual([]);
        expect(peerlessKnownState(store)).toBe(true);
        expect(peerlessInputReference(store)).toEqual({ source: 'confirmed', selection: selection(3) });
      }
      expect(parse.mock.calls.filter(([text]) => text === new TextDecoder().decode(state.frame.displayList))).toHaveLength(0);
      expect(paintFrame).toHaveBeenCalledTimes(1);
      expect(paintEcho).toHaveBeenCalledTimes(1);
      expect(restoreLine).toHaveBeenCalledTimes(1);
      expect(publications).not.toHaveBeenCalled();
    } finally { parse.mockRestore(); store.dispose(); await client.dispose(); }
  });

  for (const source of ['stale_ack', 'thumbnail'] as const) {
    test(`late_media_repaints_installed_frame_and_thumbnail_without_revisit_${mode}_${source}`, async () => {
      const state = initial();
      state.slides.push({ ...state.slides[1]!, id: 'last', index: 2 });
      const assets = [{ assetId: 'shared-image', bytes: new Uint8Array([1, 2]).buffer }];
      const picture = (slideId: string, seq: number, media: PresentationCommandFrame['media'] = []) => {
        const value = frame(slideId, seq, seq, { slideId, version: state.version, media });
        const list = JSON.parse(new TextDecoder().decode(value.displayList)) as SlideDisplayList;
        list.primitives.push({ kind: 'image', objectId: 4, name: 'Picture', assetId: 'shared-image', x: 400, y: 100, w: 100, h: 100 });
        value.displayList = new TextEncoder().encode(JSON.stringify(list)).buffer;
        return value;
      };
      const first = picture('other', 1, assets);
      const last = picture('last', 2);
      const preview = picture('other', 2);
      const projection = createPresentationCommandProjection();
      projection.initialize(state.version);
      let deliver: (message: unknown) => void = () => {};
      const held: (() => void)[] = [];
      const transport: SessionTransport = {
        post(message) {
          if (!isClientMessage(message) || message.kind !== 'call') return;
          const reply = (value: unknown) => deliver({ protocol: 1, kind: 'reply', id: message.id, ok: true, value });
          if (message.method === 'openCommandEditor') reply({ ...state, shared: projection.ack(state) });
          else if (message.method === 'command' || message.method === 'edit') {
            const command = message.args[0] as PresentationCommandEnvelope;
            const ack: PresentationCommandAck = { ...state, throughSeq: command.seq, selection: null,
              activeSlideId: command.seq === 1 ? 'other' : 'last', frame: command.seq === 1 ? first : last,
              outcomes: [{ seq: command.seq, status: 'applied', committed: false }] };
            held.push(() => reply(mode === 'shared' ? projection.ack(ack) : ack));
          } else if (message.method === 'thumbnailFrame') {
            const request = message.args[0] as { slideId: string };
            reply({ session: state.session, incarnation: state.incarnation, epoch: state.epoch,
              slideId: request.slideId, slideRevision: 2, status: 'frame',
              frame: request.slideId === 'last' ? picture('last', 2, assets) : preview });
          }
        },
        listen(listener) { deliver = listener; return () => { deliver = () => {}; }; },
        onError() { return () => {}; },
        close() { held.length = 0; },
      };
      const client = await createClient(new Uint8Array(), { session: state.session, incarnation: state.incarnation }, transport);
      const paintFrame = mock((_frame: PeerlessConfirmedFrame, _view: PeerlessViewSnapshot) => {});
      const store = createPeerlessViewStore(client, { paintFrame, paintEcho() {}, restoreLine() {} });
      const structure = mock(() => {});
      store.subscribeStructure(structure);
      try {
        void store.dispatch({ kind: 'activateSlide', payload: { slideId: 'other' } }).catch(() => {});
        void store.dispatch({ kind: 'activateSlide', payload: { slideId: 'last' } }).catch(() => {});
        expect(held).toHaveLength(2);
        held[1]!();
        for (let turn = 0; turn < 32; turn += 1) await Promise.resolve();
        expect(store.getSnapshot().confirmed?.frame.slideId).toBe('last');
        const thumbnail = await client.thumbnail('other');
        if (thumbnail.status !== 'frame') throw new Error('Expected thumbnail');
        expect(store.installThumbnail(thumbnail.frame)).toBe(true);
        const before = store.getSnapshot();
        expect(before.confirmed?.frame.media).toEqual([]);
        expect(before.thumbnails.get('other')?.frame.media).toEqual([]);
        const paintCount = paintFrame.mock.calls.length;
        const structureCount = structure.mock.calls.length;
        for (const mismatch of [
          { session: 'old-session' }, { incarnation: 'old-worker' }, { epoch: 0 },
          { frame: { ...first, incarnation: 'old-worker' } },
          { frame: { ...first, documentEpoch: 0 } },
        ]) {
          deliver({ protocol: 1, kind: 'event', name: 'ack', payload: {
            ...state, throughSeq: 1, frame: first, ...mismatch,
          } });
        }
        expect(store.getSnapshot()).toBe(before);
        expect(paintFrame.mock.calls.length).toBe(paintCount);
        if (source === 'stale_ack') held[0]!();
        else await client.thumbnail('last');
        for (let turn = 0; turn < 32; turn += 1) await Promise.resolve();
        const after = store.getSnapshot();
        expect(after.confirmed?.frame).toMatchObject({ slideId: 'last', coversThroughSeq: 2, media: assets });
        expect(after.thumbnails.get('other')?.frame.media).toEqual(assets);
        expect(after.confirmed?.displayList).toBe(before.confirmed?.displayList);
        expect(after.appliedThroughSeq).toBe(before.appliedThroughSeq);
        expect(after.presentedThroughSeq).toBe(before.presentedThroughSeq);
        expect(client.state.frame).toBe(last);
        expect(client.state.activeSlideId).toBe('last');
        expect(paintFrame.mock.calls.length).toBe(paintCount + 1);
        expect(structure.mock.calls.length).toBe(structureCount + 1);
      } finally { store.dispose(); await client.dispose(); }
    });
  }
}

test('image_frame_decodes_once_when_cached_media_is_reused', async () => {
  const state = initial();
  const list = JSON.parse(new TextDecoder().decode(state.frame.displayList)) as SlideDisplayList;
  list.primitives.push({ kind: 'image', objectId: 4, name: 'Picture', assetId: 'shared-image', x: 400, y: 100, w: 100, h: 100 });
  const displayList = JSON.stringify(list);
  state.frame.media = [{ assetId: 'shared-image', bytes: new Uint8Array([1, 2]).buffer }];
  const installed = frame('other', 1, 1, { slideId: 'other', version: state.version,
    displayList: new TextEncoder().encode(displayList).buffer });
  let deliver: (message: unknown) => void = () => {};
  const transport: SessionTransport = {
    post(message) {
      if (!isClientMessage(message) || message.kind !== 'call') return;
      const value = message.method === 'openCommandEditor' ? state : { ...state, throughSeq: 1,
        activeSlideId: 'other', selection: null, frame: installed,
        outcomes: [{ seq: 1, status: 'applied', committed: false }] };
      deliver({ protocol: 1, kind: 'reply', id: message.id, ok: true, value });
    },
    listen(listener) { deliver = listener; return () => {}; },
    onError() { return () => {}; },
    close() {},
  };
  const client = await createPresentationCommandClient(new Uint8Array(),
    { session: state.session, incarnation: state.incarnation }, transport);
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
  const parse = spyOn(JSON, 'parse');
  try {
    let completed = false;
    void store.dispatch({ kind: 'activateSlide', payload: { slideId: 'other' } }).then(
      () => { completed = true; }, () => {});
    for (let turn = 0; turn < 32; turn += 1) await Promise.resolve();
    expect(completed).toBe(true);
    expect(store.getSnapshot().confirmed?.frame.media).toEqual(state.frame.media);
    expect(parse.mock.calls.filter(([text]) => text === displayList)).toHaveLength(1);
    expect(client.state.frame).toBe(installed);
  } finally { parse.mockRestore(); store.dispose(); await client.dispose(); }
});

test('rec_refusal_drained_tickets_keep_caret_unknown_until_covered', async () => {
  const state = initial();
  const projection = createPresentationCommandProjection();
  projection.initialize(state.version);
  let deliver: (message: unknown) => void = () => {};
  let ack: PresentationCommandAck = state;
  let text = 'abc';
  const sent: PresentationSharedEnvelope[] = [];
  const held: (() => void)[] = [];
  const transport: SessionTransport = {
    post(message) {
      if (!isClientMessage(message) || message.kind !== 'call') return;
      const reply = (value: unknown) => deliver({ protocol: 1, kind: 'reply', id: message.id, ok: true, value });
      if (message.method === 'openCommandEditor') reply({ ...state, shared: projection.ack(state) });
      else if (message.method === 'edit') {
        const command = message.args[0] as PresentationSharedEnvelope;
        sent.push(command);
        const intent = command as PresentationCommandIntent;
        let outcome: PresentationCommandAck['outcomes'][number] = {
          seq: command.seq, status: 'applied', committed: false,
        };
        if (command.seq <= 256) outcome = { seq: command.seq,
          status: command.seq === 1 ? 'refused' : 'cancelled',
          reason: command.seq === 1 ? 'bad-target' : 'dependency-refused', committed: false };
        else if (intent.kind === 'setReadOnly') ack = { ...ack,
          editState: { ...ack.editState, readOnly: intent.payload.readOnly } };
        else if (intent.kind === 'moveSelection') ack = { ...ack, selection: selection(text.length) };
        else if (intent.kind === 'insertText') {
          const position = intent.payload.at.source === 'confirmed'
            ? intent.payload.at.selection.focus : ack.selection!.focus;
          text = text.slice(0, position) + intent.payload.text + text.slice(position);
          ack = { ...ack, selection: selection(position + intent.payload.text.length), version: `v${command.seq}` };
          projection.commit(ack.version);
          outcome = { ...outcome, committed: true };
        }
        ack = { ...ack, throughSeq: command.seq, outcomes: [outcome], frame: undefined };
        const value = projection.ack(ack);
        held.push(() => reply(value));
      }
    },
    listen(listener) { deliver = listener; return () => { deliver = () => {}; }; },
    onError() { return () => {}; },
    close() { held.length = 0; },
  };
  const client = await createSharedPresentationCommandClient(new Uint8Array(),
    { session: state.session, incarnation: state.incarnation, clientId: 1 }, transport);
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} },
    { echoEnabled: true });
  let completed = 0;
  const dispatch = (intent: PresentationCommandIntent) => {
    const operation = store.dispatch(intent);
    void operation.then(() => { completed += 1; }, () => {});
    return operation;
  };
  try {
    for (let index = 0; index < 256; index += 1) dispatch({ kind: 'moveSelection', payload: {
      at: { source: 'confirmed', selection: index === 0 ? { ...selection(), storyId: 'missing' } : selection() },
      motion: 'lineEnd', extend: false, inputTime: 100, undoGroup: 'navigation',
    } });
    dispatch({ kind: 'setReadOnly', payload: { readOnly: true, inputTime: 101, undoGroup: 'mode' } });
    dispatch({ kind: 'setReadOnly', payload: { readOnly: false, inputTime: 102, undoGroup: 'mode' } });
    dispatch({ kind: 'moveSelection', payload: { at: { source: 'confirmed', selection: selection() },
      motion: 'lineEnd', extend: false, inputTime: 103, undoGroup: 'navigation' } });
    dispatch(insert('X'));
    expect(sent).toHaveLength(256);
    expect(store.getSnapshot().waitingInputCount).toBe(4);
    held.shift()!();
    for (let turn = 0; turn < 16; turn += 1) await Promise.resolve();
    expect(store.getSnapshot().waitingInputCount).toBe(0);
    for (const reply of held.splice(0, 255)) reply();
    for (let turn = 0; turn < 16; turn += 1) await Promise.resolve();
    expect(sent).toHaveLength(260);
    expect(installCoveringFrame(store, frame('abc', 1, 256, { version: 'v0' }))).toBe(true);
    expect(peerlessKnownState(store)).toBe(false);
    expect(store.getSnapshot().pending.map((record) => record.seq)).toEqual([257, 258, 259, 260]);
    expect(store.getSnapshot().echo).toBeNull();
    const at = peerlessInputReference(store)!;
    expect(at).toEqual({ source: 'afterSeq', seq: 260 });
    dispatch({ kind: 'insertText', payload: { text: 'Y', at, inputTime: 104, undoGroup: 'next' } });
    expect(store.getSnapshot().waitingInputCount).toBe(1);
    expect(peerlessKnownState(store)).toBe(false);
    for (const reply of held.splice(0)) reply();
    for (let turn = 0; turn < 16; turn += 1) await Promise.resolve();
    expect(sent[260]).toMatchObject({ seq: 261, afterSeq: 260,
      payload: { at: { source: 'afterSeq', seq: 260 }, text: 'Y' } });
    expect(text).toBe('abcXY');
    for (const reply of held.splice(0)) reply();
    for (let turn = 0; turn < 16; turn += 1) await Promise.resolve();
    expect(peerlessKnownState(store)).toBe(false);
    expect(installCoveringFrame(store, frame(text, 2, 261, { version: ack.version }))).toBe(true);
    expect(peerlessKnownState(store)).toBe(true);
    expect(store.getSnapshot().selection).toEqual(selection(5));
    expect(completed).toBe(261);
  } finally { store.dispose(); await client.dispose(); projection.dispose(); }
});

test('echo_off_by_default_paints_no_typing_echo', async () => {
  const client = new FakeCommandClient();
  const paintFrame = mock((_frame: PeerlessConfirmedFrame, _view: PeerlessViewSnapshot) => {});
  const paintEcho = mock((_view: PeerlessViewSnapshot) => {});
  const store = createPeerlessViewStore(client, { paintFrame, paintEcho, restoreLine() {} });
  try {
    expect(client.state.editState.toolbar.echoRun).toBe('after');
    const frames = paintFrame.mock.calls.length;
    const pending = store.dispatch(insert('b'));
    expect(store.getSnapshot().pending).toMatchObject([{ seq: 1, echoed: false }]);
    expect(store.getSnapshot().echo).toBeNull();
    expect(store.getSnapshot().activeLine?.text).toBe('abc');
    expect(paintEcho.mock.calls.every(([view]) => view.echo === null)).toBe(true);
    expect(paintFrame).toHaveBeenCalledTimes(frames);
    client.ack(1);
    await pending;
    expect(store.getSnapshot().echo).toBeNull();
    expect(paintFrame).toHaveBeenCalledTimes(frames);
    client.deliverFrame(frame('abbc', 1, 1));
    expect(paintFrame).toHaveBeenCalledTimes(frames + 1);
    expect(store.getSnapshot().activeLine?.text).toBe('abbc');
    expect(store.getSnapshot()).toMatchObject({ echo: null, pending: [] });
    expect(paintEcho.mock.calls.every(([view]) => view.echo === null)).toBe(true);
  } finally { store.dispose(); }
});

test('echo_style_wrap_boundary_requires_the_named_rendered_neighbour', async () => {
  const client = new FakeCommandClient('ab');
  client.state.selection = { ...selection(1), affinity: 'upstream' };
  const list = JSON.parse(new TextDecoder().decode(client.state.frame.displayList)) as SlideDisplayList;
  const box = list.primitives[0]!;
  if (box.kind !== 'textBox') throw new Error('Expected text box');
  const first = box.lines[0]!;
  box.lines = [
    { ...first, start: 0, end: 1, runs: [{ ...first.runs[0]!, text: 'a', start: 0, end: 1, bold: true }],
      caretStops: first.caretStops.filter((stop) => stop.position <= 1) },
    { ...first, start: 1, end: 2, y: 60, baseline: 75,
      runs: [{ ...first.runs[0]!, text: 'b', start: 1, end: 2, italic: true }],
      caretStops: first.caretStops.filter((stop) => stop.position >= 1) },
  ];
  client.state.frame.displayList = new TextEncoder().encode(JSON.stringify(list)).buffer;
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
  try {
    const pending = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: client.state.selection! }, text: 'X', inputTime: 100, undoGroup: 'typing',
    } });
    expect(store.getSnapshot().echo).toBeNull();
    expect(client.sent).toHaveLength(1);
    client.ack(1, { frame: frame('aXb', 1, 1) });
    await pending;
  } finally { store.dispose(); }
});

test('p2_echo_retires_only_after_covering_paint', async () => {
  const { client, store, paints, paintFrame } = setup();
  const publications: PeerlessViewSnapshot[] = [];
  store.subscribePresentation(() => { publications.push(store.getSnapshot()); });
  const first = store.dispatch(insert('X'));
  const second = store.dispatch(insert('Y', 1));
  expect(store.getSnapshot().echo).toMatchObject({ text: 'aXYbc', sequences: [1, 2] });
  const beforeAck = paintFrame.mock.calls.length;
  client.ack(1);
  await first;
  expect(store.getSnapshot()).toMatchObject({ appliedThroughSeq: 1, presentedThroughSeq: 0 });
  expect(store.getSnapshot().echo?.sequences).toEqual([1, 2]);
  expect(paintFrame).toHaveBeenCalledTimes(beforeAck);

  publications.length = 0;
  const covering = frame('aXbc', 1, 1);
  client.deliverFrame(covering);
  expect(paintFrame).toHaveBeenCalledTimes(beforeAck + 1);
  expect(paints[paints.length - 2]).toMatchObject({ kind: 'frame', view: {
    confirmed: { frame: covering }, presentedThroughSeq: 1,
    echo: { text: 'aXYbc', sequences: [2] },
  } });
  expect(paints[paints.length - 1]).toMatchObject({ kind: 'echo', view: { echo: { text: 'aXYbc', sequences: [2] } } });
  expect(publications.length).toBeGreaterThan(0);
  expect(publications.every((view) => view.confirmed!.frame === covering
    && view.echo?.sequences.join(',') === '2')).toBe(true);
  client.ack(2, { frame: frame('aXYbc', 2, 2) });
  await second;
  expect(store.getSnapshot()).toMatchObject({ echo: null, pending: [], presentedThroughSeq: 2 });
  expect(client.read).not.toHaveBeenCalled();
  store.dispose();
});

test('p2_covering_frame_rebases_surviving_confirmed_echo', async () => {
  const { client, store } = setup('abc', false);
  const first = store.dispatch(insert('X'));
  const selected = store.dispatch({ kind: 'setSelection', payload: { selection: selection(3), inputTime: 100, undoGroup: 'selection' } });
  const second = store.dispatch(insert('!', 2));
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().pending[0]?.echoed).toBe(false);
  client.ack(1, { frame: frame('aXbc', 1, 1) });
  await first;
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().selection).toEqual(selection(2));
  expect(client.envelopes.map((envelope) => envelope.afterSeq)).toEqual([undefined, 1, 2]);
  client.ack(3, { selection: selection(5), frame: frame('aXbc!', 2, 3) });
  await Promise.all([selected, second]);
  expect(store.getSnapshot().activeLine?.text).toBe('aXbc!');
  expect(store.getSnapshot().echo).toBeNull();
  store.dispose();
});

test('p2_acknowledged_mutation_is_never_a_new_result_reference', async () => {
  const { client, store } = setup();
  const inserted = store.dispatch(insert('X'));
  client.ack(1);
  await inserted;
  expect(store.getSnapshot().pending[0].status).toBe('applied');
  expect(peerlessInputReference(store)).toEqual({ source: 'afterSeq', seq: 1 });
  client.deliverFrame(frame('aXbc', 1, 1));
  expect(peerlessInputReference(store)).toEqual({ source: 'confirmed', selection: selection(2) });
  expect(store.getSnapshot().activeLine?.text).toBe('aXbc');
  store.dispose();
});

test('echo_style_local_caret_waits_for_worker_confirmation', async () => {
  const { client, store } = setup();
  const first = store.dispatch({ kind: 'setSelection', payload: { selection: selection(2), inputTime: 100, undoGroup: 'selection' } });
  expect(peerlessKnownState(store)).toBe(true);
  const second = store.dispatch({ kind: 'setSelection', payload: { selection: selection(3), inputTime: 100, undoGroup: 'selection' } });
  expect(store.getSnapshot().selection).toEqual(selection(3));
  client.ack(1, { selection: selection(2), frame: frame('abc', 1, 1) });
  await first;
  expect(peerlessKnownState(store)).toBe(true);
  expect(store.getSnapshot().selection).toEqual(selection(3));
  const typed = store.dispatch(insert('!', undefined, 3));
  expect(peerlessKnownState(store)).toBe(false);
  expect(store.getSnapshot().echo).toBeNull();
  client.ack(3, { selection: selection(4), frame: frame('abc!', 2, 3) });
  await Promise.all([second, typed]);
  expect(store.getSnapshot().activeLine?.text).toBe('abc!');
  expect(peerlessKnownState(store)).toBe(true);
  store.dispose();
});

test('p4_ack_without_frame_keeps_navigation_and_clear_selection_in_the_chain', async () => {
  const { client, store } = setup();
  expect(peerlessKnownState(store)).toBe(true);
  const inserted = store.dispatch(insert('X'));
  client.ack(1);
  await inserted;
  expect(peerlessKnownState(store)).toBe(false);
  const moved = store.dispatch({ kind: 'moveSelection', payload: {
    at: { source: 'afterSeq', seq: 1 }, motion: 'left', extend: false, inputTime: 100, undoGroup: 'selection' } });
  const cleared = store.dispatch({ kind: 'setSelection', payload: { selection: null, inputTime: 100, undoGroup: 'selection' } });
  expect(client.envelopes.map((envelope) => envelope.afterSeq)).toEqual([undefined, 1, 2]);
  expect(client.envelopes.map((envelope) => envelope.basis.version)).toEqual(['v0', 'v0', 'v0']);
  expect(store.getSnapshot().selection).toEqual(selection());
  client.ack(3, { selection: null });
  await Promise.all([moved, cleared]);
  expect(peerlessKnownState(store)).toBe(false);
  client.deliverFrame(frame('aXbc', 1, 3));
  expect(peerlessKnownState(store)).toBe(true);
  expect(store.getSnapshot().selection).toBeNull();
  expect(peerlessInputReference(store)).toBeNull();
  store.dispose();
});

test('p4_capped_navigation_and_typing_retain_the_entry_basis_and_predecessors', async () => {
  const { client, store } = setup();
  const inserts = Array.from({ length: PEERLESS_PENDING_INTENTS }, (_, index) =>
    store.dispatch(insert('x', index === 0 ? undefined : index)));
  const moved = store.dispatch({ kind: 'moveSelection', payload: {
    at: { source: 'afterSeq', seq: 256 }, motion: 'left', extend: false, inputTime: 100, undoGroup: 'selection' } });
  const typed = store.dispatch(insert('Y', 257));
  expect(store.getSnapshot().waitingInputCount).toBe(2);
  client.ack(256, { selection: selection(257), frame: frame(`a${'x'.repeat(256)}bc`, 1, 256) });
  await Promise.all(inserts);
  expect(client.envelopes[256]).toMatchObject({ basis: { version: 'v0' }, afterSeq: 256,
    kind: 'moveSelection', payload: { at: { source: 'afterSeq', seq: 256 }, motion: 'left' } });
  expect(client.envelopes[257]).toMatchObject({ basis: { version: 'v0' }, afterSeq: 257,
    payload: { at: { source: 'afterSeq', seq: 257 }, text: 'Y' } });
  expect(store.getSnapshot().echo).toBeNull();
  const text = `a${'x'.repeat(255)}Yxbc`;
  client.ack(258, { selection: selection(257), frame: frame(text, 2, 258) });
  await Promise.all([moved, typed]);
  expect(store.getSnapshot().activeLine?.text).toBe(text);
  expect(peerlessKnownState(store)).toBe(true);
  store.dispose();
});

test('p4_expired_entry_basis_uses_a_successful_covered_predecessor', async () => {
  const { client, store } = setup();
  const inserts = Array.from({ length: PEERLESS_PENDING_INTENTS }, (_, index) =>
    store.dispatch(insert('x', index === 0 ? undefined : index)));
  const move = (seq: number) => store.dispatch({ kind: 'moveSelection', payload: {
    at: { source: 'afterSeq', seq }, motion: 'left', extend: false, inputTime: 100, undoGroup: 'selection' } });
  const firstMove = move(256);
  const firstType = store.dispatch(insert('Y', 257));
  const secondMove = move(258);
  const secondType = store.dispatch(insert('Z', 259));
  client.ack(256, { selection: selection(257), frame: frame(`a${'x'.repeat(256)}bc`, 256, 256) });
  await Promise.all(inserts);
  expect(client.sent).toHaveLength(258);
  expect(client.envelopes[257].basis.version).toBe('v0');
  client.ack(258, { selection: selection(257), frame: frame(`a${'x'.repeat(255)}Yxbc`, 257, 258) });
  await Promise.all([firstMove, firstType]);
  expect(client.envelopes[258]).toMatchObject({ basis: { version: 'v0' }, afterSeq: 258,
    kind: 'moveSelection', payload: { at: { source: 'afterSeq', seq: 258 } } });
  expect(client.envelopes[259]).toMatchObject({ basis: { version: 'v0' }, afterSeq: 259,
    payload: { at: { source: 'afterSeq', seq: 259 }, text: 'Z' } });
  const text = `a${'x'.repeat(255)}ZYxbc`;
  client.ack(260, { selection: selection(257), frame: frame(text, 258, 260) });
  await Promise.all([secondMove, secondType]);
  expect(store.getSnapshot().activeLine?.text).toBe(text);
  expect(peerlessKnownState(store)).toBe(true);
  store.dispose();
});

test('p2_held_typing_coalesces_waiting_operations_and_retains_all_text', async () => {
  const { client, store } = setup();
  const pending: Promise<PresentationCommandAck>[] = [];
  for (let index = 0; index < PEERLESS_PENDING_INTENTS; index += 1) {
    pending.push(store.dispatch(insert('x', index === 0 ? undefined : index)));
  }
  const queued = store.dispatch(insert('y', 256));
  for (let index = 1; index < 10_000; index += 1) {
    const at = peerlessInputReference(store);
    if (!at) throw new Error('Expected input reference');
    expect(store.dispatch({ kind: 'insertText', payload: { at, text: 'y', inputTime: 100, undoGroup: 'typing' } })).toBe(queued);
  }
  expect(store.getSnapshot().waitingInputCount).toBe(1);
  expect(store.getSnapshot().pending).toHaveLength(256);
  client.ack(256, { selection: selection(257), frame: frame(`a${'x'.repeat(256)}bc`, 1, 256) });
  await Promise.all(pending);
  expect(client.sent[256]).toMatchObject({ payload: { text: 'y'.repeat(10_000),
    at: { source: 'afterSeq', seq: 256 } } });
  const text = `a${'x'.repeat(256)}${'y'.repeat(10_000)}bc`;
  client.ack(257, { selection: selection(10_257), frame: frame(text, 2, 257) });
  await queued;
  expect(store.getSnapshot().confirmed!.displayList.primitives[0]).toMatchObject({ lines: [{ runs: [{ text }] }] });
  expect(store.getSnapshot().pending).toEqual([]);
  expect(store.getSnapshot().waitingInputCount).toBe(0);
  store.dispose();
});

test('p2_late_frame_cannot_rewind_selection_or_echo', async () => {
  const { client, store, paintFrame, paintEcho } = setup('abc', false);
  const first = store.dispatch(insert('b'));
  client.ack(1, { frame: frame('abbc', 2, 1),
    editState: { ...client.state.editState, toolbar: { ...client.state.editState.toolbar, echoRun: 'after' } } });
  await first;
  const second = store.dispatch(insert('b', undefined, 2));
  const before = store.getSnapshot();
  expect(before.echo).toBeNull();
  expect(before.pending[0]?.echoed).toBe(false);
  const frameCount = paintFrame.mock.calls.length;
  const echoCount = paintEcho.mock.calls.length;
  for (const stale of [
    frame('abc', 1, 0),
    frame('old', 99, 99, { incarnation: 'old-worker' }),
    frame('old', 99, 99, { documentEpoch: 0 }),
    frame('other', 99, 99, { slideId: 'other', slideIndex: 1 }),
  ]) {
    expect(installCoveringFrame(store, stale)).toBe(false);
    expect(store.getSnapshot()).toBe(before);
  }
  client.ack(0, { selection: selection(0) });
  client.ack(99, { session: 'other-session', incarnation: 'old-worker', selection: selection(0) });
  expect(store.getSnapshot()).toBe(before);
  expect(store.getSnapshot().selection).toMatchObject({ anchor: 2, focus: 2 });
  expect(store.getSnapshot().echo).toBeNull();
  expect(paintFrame).toHaveBeenCalledTimes(frameCount);
  expect(paintEcho).toHaveBeenCalledTimes(echoCount);
  client.ack(2, { frame: frame('abbbc', 3, 2) });
  await second;
  store.dispose();
});

test('echo_allowlist_cap_never_echoes_parked_or_merged_chunks', async () => {
  const { client, store, paints } = setup();
  const accepted = Array.from({ length: PEERLESS_PENDING_INTENTS }, (_, index) =>
    store.dispatch(insert('a', index === 0 ? undefined : index)));
  expect(store.getSnapshot().pending.every((record) => record.echoed)).toBe(true);
  expect(store.getSnapshot().echo?.sequences).toHaveLength(256);
  const parked = store.dispatch(insert('b', 256));
  const newline = store.dispatch(insert('\nY', 257));
  const merged = store.dispatch(insert('a', 257));
  expect(newline).toBe(parked);
  expect(merged).toBe(parked);
  const further = store.dispatch({ kind: 'insertText', payload: {
    at: { source: 'afterSeq', seq: 257 }, text: 'a', inputTime: 101, undoGroup: 'later',
  } });
  expect(store.getSnapshot().waitingInputCount).toBe(2);
  expect(store.getSnapshot().echo?.sequences).not.toContain(257);
  expect(store.getSnapshot().echo?.sequences).not.toContain(258);
  const covered = `a${'a'.repeat(256)}bc`;
  client.ack(256, { selection: selection(257), frame: frame(covered, 1, 256) });
  await Promise.all(accepted);
  expect(client.sent[256]).toMatchObject({ payload: { text: 'b\nYa' } });
  expect(store.getSnapshot().pending).toMatchObject([{ seq: 257, echoed: false }]);
  expect(store.getSnapshot().echo).toBeNull();
  expect(paints.filter((paint) => paint.view.echo?.sequences.includes(257))).toEqual([]);
  const withNewline = `a${'a'.repeat(256)}b\nYabc`;
  client.ack(257, { selection: selection(261), frame: frame(withNewline, 2, 257) });
  await Promise.all([parked, newline, merged]);
  expect(store.getSnapshot().activeLine?.text).toBe(withNewline);
  expect(store.getSnapshot().pending).toMatchObject([{ seq: 258, echoed: false }]);
  expect(store.getSnapshot().echo).toBeNull();
  expect(paints.filter((paint) => paint.view.echo?.sequences.includes(258))).toEqual([]);
  const text = `a${'a'.repeat(256)}b\nYaabc`;
  client.ack(258, { selection: selection(262), frame: frame(text, 3, 258) });
  await further;
  expect(store.getSnapshot().activeLine?.text).toBe(text);
  store.dispose();
});

test('echo_allowlist_touching_fallback_run_suppresses_latin_key', async () => {
  for (const [text, fontId] of [['ش', 1], ['b', 1], ['ش', 0]] as const) {
    const client = new FakeCommandClient();
    const list = JSON.parse(new TextDecoder().decode(client.state.frame.displayList)) as SlideDisplayList;
    const box = list.primitives[0]!;
    if (box.kind !== 'textBox') throw new Error('Expected text box');
    const run = box.lines[0]!.runs[0]!;
    box.lines[0]!.runs = [{ ...run, text: 'a', start: 0, end: 1 },
      { ...run, text, fontId, start: 1, end: 2 }, { ...run, text: 'c', start: 2, end: 3 }];
    client.state.selection = selection(1);
    client.state.editState.toolbar.echoRun = 'before';
    client.state.frame.displayList = new TextEncoder().encode(JSON.stringify(list)).buffer;
    const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
    try {
      const pending = store.dispatch(insert('a'));
      expect(store.getSnapshot().echo).toBeNull();
      expect(store.getSnapshot().pending[0]?.echoed).toBe(false);
      client.ack(1, { frame: frame(`aa${text}c`, 1, 1) });
      await pending;
      expect(store.getSnapshot().activeLine?.text).toBe(`aa${text}c`);
    } finally { store.dispose(); }
  }
});

test('echo_allowlist_budget_parks_following_covered_key', async () => {
  const text = 'a'.repeat(PEERLESS_ACTIVE_CONTEXT_BYTES / 2);
  const { client, store, paints } = setup(text);
  const first = store.dispatch(insert('a'));
  expect(store.getSnapshot().pending[0]?.echoed).toBe(false);
  expect(store.getSnapshot().echo).toBeNull();
  const second = store.dispatch(insert('a', 1));
  expect(store.getSnapshot().waitingInputCount).toBe(1);
  expect(store.getSnapshot().echo).toBeNull();
  client.ack(1, { frame: frame(`${text}a`, 1, 1) });
  await first;
  expect(store.getSnapshot().pending).toMatchObject([{ seq: 2, echoed: false }]);
  expect(store.getSnapshot().echo).toBeNull();
  client.ack(2, { frame: frame(`${text}aa`, 2, 2) });
  await second;
  expect(store.getSnapshot().confirmed!.displayList.primitives[0]).toMatchObject({ lines: [{ runs: [{ text: `${text}aa` }] }] });
  expect(paints.filter((paint) => paint.view.echo?.sequences.includes(1) || paint.view.echo?.sequences.includes(2))).toEqual([]);
  store.dispose();
});

test('p2_refusal_drops_dependents_and_restores_line', async () => {
  const { client, store, restoreLine, paints } = setup('abc', false);
  const first = store.dispatch(insert('X'));
  const dependent = store.dispatch(insert('Y', 1));
  const later = store.dispatch(insert('!', undefined, 3));
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().pending[0]?.echoed).toBe(false);
  expect(client.envelopes.map((envelope) => envelope.afterSeq)).toEqual([undefined]);
  expect(store.getSnapshot().pending.map((record) => record.afterSeq)).toEqual([undefined]);
  expect(store.getSnapshot().waitingInputCount).toBe(1);
  expect(later).toBe(dependent);
  const restores = restoreLine.mock.calls.length;
  client.ack(1, { version: 'v0', selection: selection(1), outcomes: [
    { seq: 1, status: 'refused', reason: 'read-only', committed: false },
  ], dirtySlideIds: [] });
  await Promise.all([first, dependent]);
  expect(store.getSnapshot().pending).toEqual([]);
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().selection).toMatchObject({ anchor: 1, focus: 1 });
  expect(restoreLine).toHaveBeenCalledTimes(restores + 1);
  expect(paints[paints.length - 2]?.kind).toBe('restore');
  expect(paints[paints.length - 1]).toMatchObject({ kind: 'echo', view: { echo: null } });
  expect((await later).outcomes).toEqual([{ seq: 2, status: 'cancelled', reason: 'dependency-refused', committed: false }]);
  client.deliverFrame(frame('abc', 1, 1));
  store.setSelection(selection(3));
  const independent = store.dispatch(insert('!', undefined, 3));
  client.ack(2, { selection: selection(4), frame: frame('abc!', 2, 2) });
  await independent;
  const refused = store.dispatch(insert('?', undefined, 4));
  client.ack(3, { version: 'v2', selection: selection(4), dirtySlideIds: [],
    outcomes: [{ seq: 3, status: 'refused', reason: 'read-only', committed: false }], frame: frame('abc!', 3, 3) });
  await refused;
  expect(store.getSnapshot()).toMatchObject({ echo: null, pending: [], selection: selection(4) });
  expect(store.getSnapshot().activeLine?.text).toBe('abc!');
  store.dispose();
});

test('p2_refused_mutation_cancels_waiting_text_and_preserves_independent_input', async () => {
  const text = 'a'.repeat(8192);
  const { client, store } = setup(text);
  const oversized = store.dispatch(insert('BIG'));
  const dependent = store.dispatch(insert('NEXT', 1));
  const later = store.dispatch(insert('!', undefined, 0));
  client.ack(1, { version: 'v0', selection: selection(1), dirtySlideIds: [],
    outcomes: [{ seq: 1, status: 'refused', reason: 'invalid-offset', committed: false }],
    frame: frame(text, 1, 1) });
  await oversized;
  expect((await dependent).outcomes).toEqual([{ seq: 2, status: 'cancelled', reason: 'dependency-refused', committed: false }]);
  expect(later).toBe(dependent);
  expect((await later).outcomes).toEqual([{ seq: 2, status: 'cancelled', reason: 'dependency-refused', committed: false }]);
  expect(client.envelopes).toHaveLength(1);
  store.setSelection(selection(0));
  const independent = store.dispatch(insert('!', undefined, 0));
  expect(client.envelopes[1]).toMatchObject({ seq: 2, payload: { text: '!' } });
  expect(store.getSnapshot().waitingInputCount).toBe(0);
  client.ack(2, { selection: selection(1), outcomes: [{ seq: 2, status: 'applied', committed: true }],
    frame: frame(`!${text}`, 2, 2) });
  await independent;
  await store.fence();
  expect(store.getSnapshot().confirmed!.displayList.primitives[0]).toMatchObject({ lines: [{ runs: [{ text: `!${text}` }] }] });
  expect(store.getSnapshot().pending).toEqual([]);
  store.dispose();
});

test('echo_style_other_story_caret_waits_for_worker_confirmation_after_refusal', async () => {
  const { client, store } = setup('abc', false);
  const refused = store.dispatch(insert('X'));
  const dependent = store.dispatch(insert('Y', 1));
  const transitive = store.dispatch(insert('Z', 2));
  const other = { ...selection(), shapeId: 'other-box', storyId: 'other-story' };
  store.setSelection(other);
  const at = peerlessInputReference(store);
  if (!at) throw new Error('Expected input reference');
  expect(at).toEqual({ source: 'afterSeq', seq: 3 });
  const later = store.dispatch({ kind: 'insertText', payload: {
    at, text: '!', inputTime: 100, undoGroup: 'other',
  } });
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().pending[0]?.echoed).toBe(false);
  expect(transitive).toBe(dependent);
  expect(client.envelopes.map((envelope) => envelope.afterSeq)).toEqual([undefined]);
  expect(peerlessInputReference(store)).toEqual({ source: 'afterSeq', seq: 4 });
  client.ack(1, { version: 'v0', selection: selection(), dirtySlideIds: [],
    outcomes: [{ seq: 1, status: 'refused', reason: 'invalid-offset', committed: false }] });
  await refused;
  expect(store.getSnapshot().pending).toEqual([]);
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().selection).toEqual(selection());
  client.deliverFrame(frame('abc', 1, 1));
  await Promise.all([dependent, transitive, later]);
  store.setSelection(other);
  const independent = store.dispatch({ kind: 'insertText', payload: {
    at: { source: 'confirmed', selection: other }, text: '!', inputTime: 100, undoGroup: 'other' } });
  expect(store.getSnapshot().echo).toBeNull();
  client.ack(2, { selection: { ...other, anchor: 2, focus: 2 }, frame: frame('a!bc', 2, 2) });
  await independent;
  store.dispose();
});

test('p2_explicit_selection_preserves_confirmed_reference_until_typing', async () => {
  const { client, store } = setup('abc', false);
  const first = store.dispatch(insert('X'));
  const other = { ...selection(), shapeId: 'other-box', storyId: 'other-story' };
  const second = store.dispatch({ kind: 'insertText', payload: { text: '!',
    at: { source: 'confirmed', selection: other }, inputTime: 100, undoGroup: 'other' } });
  const mode = store.dispatch({ kind: 'setReadOnly', payload: { readOnly: false, inputTime: 100, undoGroup: 'mode' } });
  const selected = store.dispatch({ kind: 'setSelection', payload: { selection: selection(2),
    inputTime: 100, undoGroup: 'selection' } });
  expect(peerlessInputReference(store)).toEqual({ source: 'afterSeq', seq: 4 });
  const at = peerlessInputReference(store);
  if (!at) throw new Error('Expected input reference');
  const typed = store.dispatch({ kind: 'insertText', payload: { at, text: 'Y', inputTime: 100, undoGroup: 'typing' } });
  expect(client.envelopes.map((envelope) => envelope.afterSeq)).toEqual([undefined]);
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().pending[0]?.echoed).toBe(false);
  expect(store.getSnapshot().waitingInputCount).toBe(4);
  expect(peerlessInputReference(store)).toEqual({ source: 'afterSeq', seq: 5 });
  client.ack(1, { frame: frame('aXbc', 1, 1) });
  await first;
  client.ack(2, { frame: frame('aX!bc', 2, 2) });
  await second;
  expect(client.envelopes.map((envelope) => envelope.afterSeq)).toEqual([undefined, 1, undefined, 3, 4]);
  expect(store.getSnapshot().echo).toBeNull();
  client.ack(5);
  await Promise.all([first, second, mode, selected, typed]);
  store.dispose();
});

test('p2_terminal_failure_rejects_queued_input_and_ignores_late_ack', async () => {
  const { client, store } = setup('a'.repeat(PEERLESS_ACTIVE_CONTEXT_BYTES / 2));
  const oversized = store.dispatch(insert('BIG'));
  const waiting = store.dispatch(insert('NEXT', 1));
  const error = new CommandOutcomeUnknown(1, 'command outcome unknown');
  const rejected = [oversized, waiting].map((operation) => operation.then(
    () => { throw new Error('Expected terminal failure'); }, (reason: unknown) => reason));
  expect(store.getSnapshot().waitingInputCount).toBe(1);
  client.fail(error);
  expect(store.getSnapshot().editState.readOnly).toBe(true);
  expect(await Promise.all(rejected)).toEqual([error, error]);
  expect(store.getSnapshot().editState.readOnly).toBe(true);
  expect(store.getSnapshot().waitingInputCount).toBe(0);
  await expect(store.dispatch(insert('late'))).rejects.toBe(error);
  const last = store.getSnapshot();
  client.ack(1, { frame: frame('unexpected', 1, 1) });
  expect(store.getSnapshot()).toBe(last);
  store.dispose();
  expect(client.failureListeners.size).toBe(0);
});

test('terminal_notification_settles_input_without_a_transport_reply', async () => {
  const { client, store } = setup('a'.repeat(PEERLESS_ACTIVE_CONTEXT_BYTES / 2));
  const sent = store.dispatch(insert('BIG'));
  const queued = store.dispatch(insert('NEXT', 1));
  const rejected = [sent, queued].map((operation) => operation.then(
    () => { throw new Error('Expected terminal failure'); }, (reason: unknown) => reason));
  const error = new CommandOutcomeUnknown(1, 'worker failed without a reply');
  for (const listener of client.failureListeners) listener(error);
  expect(await Promise.all(rejected)).toEqual([error, error]);
  expect(store.getSnapshot()).toMatchObject({ waitingInputCount: 0, editState: { readOnly: true } });
  expect(client.replies.size).toBe(1);
  store.dispose();
});

test('p2_unchanged_thumbnail_survives_text_ack', async () => {
  const { client, store } = setup();
  expect(store.installThumbnail(frame('other', 7, 0, { slideId: 'other', slideIndex: 1 }))).toBe(true);
  const cached = store.getSnapshot().thumbnails.get('other');
  expect(cached).toBeDefined();
  expect(cached?.frame.slideRevision).toBe(7);
  const structural = mock(() => {});
  store.subscribeStructure(structural);
  const pending = store.dispatch(insert('X'));
  client.ack(1);
  await pending;
  expect(store.getSnapshot().thumbnails.get('other')).toBe(cached);
  expect(store.getSnapshot().dirtyThumbnailIds.has('other')).toBe(false);
  expect(client.thumbnail).not.toHaveBeenCalled();
  expect(structural).not.toHaveBeenCalled();
  client.deliverFrame(frame('aXbc', 1, 1));
  expect(store.getSnapshot().thumbnails.get('other')).toBe(cached);
  expect(structural).not.toHaveBeenCalled();
  store.dispose();
});

test('p2_echo_caps_wait_instead_of_dropping', async () => {
  const { client, store } = setup();
  const accepted: Promise<PresentationCommandAck>[] = [];
  for (let index = 0; index < PEERLESS_PENDING_INTENTS; index += 1) {
    accepted.push(store.dispatch(insert('x', index === 0 ? undefined : index)));
  }
  expect(store.getSnapshot().pending).toHaveLength(256);
  expect(client.sent).toHaveLength(256);
  const cappedEcho = store.getSnapshot().echo;
  let overflowSettled = false;
  const overflow = store.dispatch(insert('LAST', 256));
  void overflow.then(() => { overflowSettled = true; });
  expect(store.getSnapshot().waitingInputCount).toBe(1);
  expect(store.getSnapshot().pending).toHaveLength(256);
  expect(store.getSnapshot().echo).toBe(cappedEcho);
  expect(client.sent).toHaveLength(256);
  expect(overflowSettled).toBe(false);
  client.ack(256, { selection: selection(257), frame: frame(`a${'x'.repeat(256)}bc`, 1, 256) });
  await Promise.all(accepted);
  await Promise.resolve();
  expect(client.sent).toHaveLength(257);
  expect(client.sent[256]).toMatchObject({ kind: 'insertText', payload: { text: 'LAST' } });
  expect(store.getSnapshot().waitingInputCount).toBe(0);
  client.ack(257, { selection: selection(261), frame: frame(`a${'x'.repeat(256)}LASTbc`, 2, 257) });
  await overflow;
  expect(overflowSettled).toBe(true);
  store.dispose();

  const large = setup('a'.repeat(PEERLESS_ACTIVE_CONTEXT_BYTES / 2));
  const before = large.store.getSnapshot();
  const oversized = large.store.dispatch(insert('BIG', undefined, 1));
  expect(large.store.getSnapshot().activeContextBytes).toBeLessThanOrEqual(16 * 1024);
  expect(large.store.getSnapshot().echo).toBe(before.echo);
  expect(large.client.sent).toHaveLength(1);
  expect(large.store.getSnapshot().pending[0]).toMatchObject({ seq: 1, echoed: false });
  const waiting = large.store.dispatch(insert('NEXT', 1));
  expect(large.store.getSnapshot().waitingInputCount).toBe(1);
  expect(large.client.sent).toHaveLength(1);
  large.client.ack(1, { selection: selection(4), frame: frame(`aBIG${'a'.repeat(8191)}`, 1, 1) });
  await oversized;
  await Promise.resolve();
  expect(large.client.sent).toHaveLength(2);
  expect(large.store.getSnapshot().activeContextBytes).toBeLessThanOrEqual(16 * 1024);
  expect(large.client.sent[1]).toMatchObject({ kind: 'insertText', payload: { text: 'NEXT' } });
  large.client.ack(2, { selection: selection(8), frame: frame(`aBIGNEXT${'a'.repeat(8191)}`, 2, 2) });
  await waiting;
  large.store.dispose();
});

test('confirmed geometry preserves the screen-sized border band inside table cells', () => {
  const client = new FakeCommandClient();
  const list = JSON.parse(new TextDecoder().decode(client.state.frame.displayList)) as SlideDisplayList;
  const box = list.primitives[0]!;
  list.primitives = [{ kind: 'table', objectId: 1, shapeId: 'box', name: 'Table',
    x: 20, y: 40, w: 300, h: 40, label: 'Table', primitives: [box] }];
  client.state.frame.displayList = new TextEncoder().encode(JSON.stringify(list)).buffer;
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
  try {
    expect(hitConfirmedGeometry(store, { x: 31, y: 50 }, 0.5)?.kind).toBe('shape');
    expect(hitConfirmedGeometry(store, { x: 33, y: 53 }, 0.5)).toMatchObject({
      kind: 'text', shapeId: 'box', storyId: 'story', position: 1,
    });
    expect(hitConfirmedGeometry(store, { x: 21.4, y: 50 }, 4)?.kind).toBe('shape');
    expect(hitConfirmedGeometry(store, { x: 21.6, y: 50 }, 4)?.kind).toBe('text');
  } finally { store.dispose(); }
});

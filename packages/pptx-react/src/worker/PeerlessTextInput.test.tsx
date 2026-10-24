import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { useRef, useSyncExternalStore } from 'react';
import { CommandOutcomeUnknown } from '@betteroffice/pptx/worker';
import type { PptxReadRequest, SlideDisplayList, TextBoxPrimitive } from '@betteroffice/pptx';
import type {
  PresentationCommandAck, PresentationCommandClient, PresentationCommandFrame, PresentationCommandFailure,
  PresentationCommandIntent, PresentationCommandEnvelope, PresentationCommandOpenResult,
  PresentationCommandReadResult, PresentationCommandSelection, PresentationCommandThumbnailResult,
} from '@betteroffice/pptx/worker';
import { PeerlessTextInput, type PeerlessInputShortcut } from './PeerlessTextInput';
import { usePeerlessTextSelection } from './usePeerlessTextSelection';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import {
  createPeerlessViewStore,
  hitConfirmedGeometry,
  type PeerlessConfirmedFrame,
  type PeerlessViewSnapshot,
  type PeerlessViewStore,
} from './peerlessViewStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const disposers: (() => void)[] = [];
let mainWasmCalls = 0;
beforeEach(() => {
  mainWasmCalls = 0;
  const forbidden = () => { mainWasmCalls += 1; throw new Error('main wasm init'); };
  const spies = [
    spyOn(WebAssembly, 'instantiate').mockImplementation(forbidden),
    spyOn(WebAssembly, 'instantiateStreaming').mockImplementation(forbidden),
    spyOn(WebAssembly, 'compile').mockImplementation(forbidden),
    spyOn(WebAssembly, 'compileStreaming').mockImplementation(forbidden),
  ];
  disposers.push(...spies.map((spy) => () => { spy.mockRestore(); }));
});
afterEach(() => {
  cleanup();
  for (const dispose of disposers.splice(0).reverse()) dispose();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function selection(anchor = 1, focus = anchor): PresentationCommandSelection {
  return { slideId: 'slide', shapeId: 'box', storyId: 'story',
    anchor, focus, affinity: 'downstream' };
}

function frame(text = 'abc', revision = 0, coversThroughSeq = 0): PresentationCommandFrame {
  const caretStops = [{ position: 0, x: 20 }];
  let position = 0;
  for (const character of text) {
    position += character.length;
    caretStops.push({ position, x: 20 + position * 10 });
  }
  const box: TextBoxPrimitive = {
    kind: 'textBox', objectId: 1, shapeId: 'box', storyId: 'story',
    x: 20, y: 40, w: 300, h: 40, anchor: 'top',
    paragraphs: [{ level: 0, runs: [{ text, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
    lines: [{ x: 20, y: 40, width: text.length * 10, height: 20, baseline: 55,
      start: 0, end: text.length,
      caretStops,
      runs: [{ text, start: 0, end: text.length, x: 20, width: text.length * 10,
        fontId: 0, fontFamily: 'Arial', fontSizePx: 16, bold: false, italic: false,
        underline: false, color: '#000000', glyphs: [] }] }],
  };
  const displayList: SlideDisplayList = {
    contractVersion: 1, width: 960, height: 540,
    background: { kind: 'solid', color: '#ffffff' }, primitives: [box],
  };
  return { slideId: 'slide', slideIndex: 0, slideRevision: revision, coversThroughSeq,
    documentEpoch: 1, incarnation: 'worker', epoch: revision + 1, version: `v${revision}`,
    displayList: new TextEncoder().encode(JSON.stringify(displayList)).buffer, media: [] };
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
  private current: PresentationCommandOpenResult & { frame: PresentationCommandFrame };
  readonly sent: PresentationCommandIntent[] = [];
  readonly envelopes: PresentationCommandEnvelope[] = [];
  readonly replies = new Map<number, ReturnType<typeof deferred<PresentationCommandAck>>>();
  readonly ackListeners = new Set<(ack: PresentationCommandAck) => void>();
  readonly frameListeners = new Set<(frame: PresentationCommandFrame) => void>();
  readonly failureListeners = new Set<(error: PresentationCommandFailure) => void>();

  constructor(text = 'abc', object = false) {
    this.current = initial(text);
    if (object) this.current.selection = { ...selection(), kind: 'object', token: 'object',
      bounds: { x: 20, y: 40, width: 300, height: 40 },
      rect: { x: 20, y: 40, width: 300, height: 40 }, canMove: true, canResize: true, units: { x: 1, y: 1 } };
  }

  get state() { return this.current; }

  dispatch = mock((intent: PresentationCommandIntent, basis?: { epoch: number; version: string }): Promise<PresentationCommandAck> => {
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
    this.failureListeners.add(listener);
    return () => { this.failureListeners.delete(listener); };
  }

  ack(throughSeq: number, overrides: Partial<PresentationCommandAck> = {}) {
    const { frame: _frame, slides: _slides, size: _size, ...base } = this.state;
    const commands = this.sent.slice(0, throughSeq);
    const ack: PresentationCommandAck = { ...base, throughSeq,
      outcomes: commands.map((intent, index): PresentationCommandAck['outcomes'][number] => ({ seq: index + 1,
        status: intent.kind.startsWith('composition') && intent.kind !== 'compositionCommit' ? 'noop' : 'applied',
        committed: ['compositionCommit', 'replaceText', 'insertText', 'deleteBackward',
          'deleteForward', 'paragraphBreak'].includes(intent.kind) })),
      ...overrides };
    this.current = { ...this.current, ...ack, frame: ack.frame ?? this.current.frame };
    for (const listener of this.ackListeners) listener(ack);
    if (ack.frame) this.deliverFrame(ack.frame);
    for (const [seq, reply] of this.replies) {
      if (seq <= throughSeq) { reply.resolve(ack); this.replies.delete(seq); }
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

function setup(text = 'abc', object = false) {
  const client = new FakeCommandClient(text, object);
  const paintFrame = mock((_frame: PeerlessConfirmedFrame, _view: PeerlessViewSnapshot) => {});
  const paintEcho = mock((_view: PeerlessViewSnapshot) => {});
  const restoreLine = mock((_view: PeerlessViewSnapshot) => {});
  const store = createPeerlessViewStore(client, { paintFrame, paintEcho, restoreLine });
  disposers.push(() => { store.dispose(); });
  return { client, store, paintFrame, paintEcho, restoreLine };
}

test('refused typing remains visible even without an error callback', async () => {
  const { client, store } = setup();
  const view = render(<PeerlessTextInput store={store} />);
  await act(async () => {
    beforeInput(textarea(view.container), 'insertText', 'NEXT');
    client.ack(1, { outcomes: [{ seq: 1, status: 'refused', reason: 'stale-basis', committed: false }] });
  });
  expect(view.getByRole('alert').textContent).toContain('stale-basis');
});

function beforeInput(
  input: HTMLTextAreaElement,
  inputType: string,
  data: string | null = null,
  extra: { isComposing?: boolean; dataTransfer?: Pick<DataTransfer, 'getData' | 'types'> } = {},
) {
  const event = new Event('beforeinput', { bubbles: true, cancelable: true });
  Object.defineProperties(event, {
    inputType: { value: inputType }, data: { value: data },
    isComposing: { value: extra.isComposing ?? false }, dataTransfer: { value: extra.dataTransfer ?? null },
  });
  fireEvent(input, event);
  return event;
}

function textarea(container: HTMLElement): HTMLTextAreaElement {
  const input = container.querySelector('textarea');
  expect(input === null).toBe(false);
  if (!input) throw new Error('expected peer-less textarea');
  return input;
}

function compose(input: HTMLTextAreaElement, type: string, data: string): void {
  const event = new Event(type, { bubbles: true });
  Object.defineProperty(event, 'data', { value: data });
  fireEvent(input, event);
}

function expectNoMainModel(client: FakeCommandClient) {
  expect(client.read).not.toHaveBeenCalled();
  expect(mainWasmCalls).toBe(0);
  for (const file of ['peerlessViewStore.ts', 'PeerlessTextInput.tsx']) {
    const source = readFileSync(new URL(file, import.meta.url), 'utf8');
    expect(source).not.toMatch(/\b(?:PptxDocument|initWasm|openPresentation|attachPeer|hydrate)\b/);
    expect(source).not.toMatch(/(?:from\s*|import\s*\()["'][^"']*(?:wasm|editorSession|editPeer|commandClient|commandHost)/);
    expect(source).not.toMatch(/\b(?:story|snapshot|layoutSlide)\s*\(/);
  }
}

test('composition_after_pending_first_selection', () => {
  const { client, store } = setup();
  store.setSelection(null);
  const { container } = render(<PeerlessTextInput store={store} />);
  const input = textarea(container);
  act(() => {
    void store.dispatch({ kind: 'selectAtPoint', payload: {
      frame: client.state.frame, x: 30, y: 50, extend: false,
      inputTime: 1, undoGroup: 'pointer',
    } }).catch(() => {});
    compose(input, 'compositionstart', '');
    compose(input, 'compositionupdate', '漢');
    compose(input, 'compositionend', '漢');
  });
  expect(client.sent.map((intent) => intent.kind)).toEqual([
    'selectAtPoint', 'compositionStart', 'compositionCommit',
  ]);
  expect(client.sent[1]).toMatchObject({ payload: { at: { source: 'afterSeq', seq: 1 } } });
  expect(client.sent[2]).toMatchObject({ payload: { at: { source: 'afterSeq', seq: 2 }, text: '漢' } });
  act(() => { client.ack(3, { selection: selection(2), frame: frame('a漢bc', 1, 3) }); });
  expect(client.sent[3]).toMatchObject({ kind: 'compositionEnd',
    payload: { at: { source: 'afterSeq', seq: 3 } } });
  act(() => { client.ack(4, { selection: selection(2), frame: frame('a漢bc', 2, 4) }); });
  expect(store.getSnapshot().composition).toBeNull();
});

function TextClick({ store, presentation }: { store: PeerlessViewStore; presentation: PeerlessPresentation }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const gestures = usePeerlessTextSelection(store, { canvas, presentation });
  return <div onPointerDown={gestures.onPointerDown} onPointerUp={gestures.onPointerUp}>
    <canvas ref={canvas} /><PeerlessTextInput store={store} />
  </div>;
}

for (const kind of ['typing', 'paste', 'composition'] as const) {
  test(`selected_object_pending_text_click_${kind}`, async () => {
    const { client, store } = setup('abc', true);
    const presentation = { selectionEnabled: true, store, visual: store, require: () => {},
      getVisualSnapshot: () => ({ view: store.getSnapshot(), fit: { scale: 1, dpr: 1, generation: 1 } }),
    } as unknown as PeerlessPresentation;
    const view = render(<TextClick store={store} presentation={presentation} />);
    const canvas = view.container.querySelector('canvas')!;
    const bounds = spyOn(canvas, 'getBoundingClientRect').mockReturnValue({ left: 0, top: 0,
      right: 960, bottom: 540, width: 960, height: 540 } as DOMRect);
    disposers.push(() => bounds.mockRestore());
    const input = textarea(view.container);
    await act(async () => {
      fireEvent.pointerDown(canvas, { clientX: 40, clientY: 50, button: 0, pointerId: 1, isPrimary: true });
      if (kind === 'typing') beforeInput(input, 'insertText', '!');
      else if (kind === 'paste') fireEvent.paste(input, { clipboardData: { getData: () => 'paste' } });
      else { compose(input, 'compositionstart', ''); compose(input, 'compositionupdate', '漢'); }
    });
    expect(client.sent[0]?.kind).toBe('selectAtPoint');
    expect(client.envelopes[1]).toMatchObject({
      kind: kind === 'typing' ? 'insertText' : kind === 'paste' ? 'replaceText' : 'compositionStart',
      afterSeq: 1, payload: { at: { source: 'afterSeq', seq: 1 } },
    });
    if (kind === 'composition') expect(store.getSnapshot().composition?.text).toBe('漢');
    await act(async () => {
      client.ack(2, { slides: undefined, selection: selection(2), frame: frame('abc', 1, 2) });
      if (kind === 'composition') compose(input, 'compositionend', '漢');
    });
    if (kind === 'composition') {
      expect(client.sent.filter((intent) => intent.kind === 'compositionCommit')).toHaveLength(1);
      await act(async () => { client.ack(3, { slides: undefined, selection: selection(3), frame: frame('ab漢c', 2, 3) }); });
      await act(async () => { client.ack(4); });
    }
    expectNoMainModel(client);
  });
}

test('selected_object_without_pending_click_blocks_text_and_composition', async () => {
  const { client, store } = setup('abc', true);
  const view = render(<PeerlessTextInput store={store} />);
  const input = textarea(view.container);
  await act(async () => {
    beforeInput(input, 'insertText', '!');
    fireEvent.paste(input, { clipboardData: { getData: () => 'paste' } });
    compose(input, 'compositionstart', '');
  });
  expect(client.sent).toHaveLength(0);
  expect(store.getSnapshot().composition).toBeNull();
});

test('p2_confirmed_geometry_drives_caret_without_story_reads', async () => {
  const { client, store } = setup('a🙂b');
  const hit = hitConfirmedGeometry(store, { x: 51, y: 48 });
  expect(hit).toEqual({ kind: 'text', shapeId: 'box', storyId: 'story', position: 3, affinity: 'downstream' });
  expect(hitConfirmedGeometry(store, { x: 900, y: 400 })).toBeNull();
  store.setSelection(selection(3));
  expect(store.getSnapshot().caret).toEqual({ x: 50, y: 40, width: 1, height: 20 });
  store.setSelection(selection(0, 3));
  expect(store.getSnapshot().selectionRects).toEqual([{ x: 20, y: 40, width: 30, height: 20 }]);
  store.setSelection(selection(3));
  const shortcuts = mock((_shortcut: PeerlessInputShortcut) => {});
  const view = render(<PeerlessTextInput store={store} viewport={{ left: 100, top: 50, scale: 2 }}
    onShortcut={shortcuts} />);
  const input = textarea(view.container);
  expect(input.style.position).toBe('absolute');
  expect(input.style.left).toBe('200px');
  expect(input.style.top).toBe('130px');
  expect(input.style.opacity).toBe('0');
  expect(input.hidden).toBe(false);
  expect(input.style.display).not.toBe('none');
  expect(input.style.visibility).not.toBe('hidden');
  act(() => { input.focus(); });
  expect(document.activeElement === input).toBe(true);
  expect(store.getSnapshot().focused).toBe(true);
  await act(async () => { fireEvent.keyDown(input, { key: 'ArrowLeft' }); });
  expect(store.getSnapshot().selection).toMatchObject({ anchor: 1, focus: 1 });
  expect(client.sent[client.sent.length - 1]).toMatchObject({ kind: 'setSelection', payload: { selection: selection(1) } });
  expect(input.style.left).toBe('160px');
  await act(async () => { fireEvent.keyDown(input, { key: 'z', ctrlKey: true }); });
  expect(shortcuts).toHaveBeenCalledWith('undo');
  await act(async () => { beforeInput(input, 'deleteContentBackward'); });
  expect(client.sent[client.sent.length - 1]).toMatchObject({ kind: 'deleteBackward', payload: {
    unit: 'codePoint', at: { source: 'afterSeq', seq: 1 } } });
  await act(async () => { beforeInput(input, 'deleteContentForward'); });
  expect(store.getSnapshot().waitingInputCount).toBe(1);
  await act(async () => { client.ack(client.sent.length, { version: 'v2', selection: selection(0) }); });
  expect(client.sent[client.sent.length - 1]).toMatchObject({ kind: 'deleteForward', payload: { unit: 'codePoint', at: { source: 'afterSeq', seq: 2 } } });
  expectNoMainModel(client);
});

test('p2_composition_commits_once_and_survives_frame_install', async () => {
  const { client, store, paintEcho } = setup();
  const view = render(<PeerlessTextInput store={store} />);
  const input = textarea(view.container);
  await act(async () => {
    input.focus();
    compose(input, 'compositionstart', '');
    compose(input, 'compositionupdate', '漢');
    input.value = '漢';
    beforeInput(input, 'insertCompositionText', '漢', { isComposing: true });
    fireEvent.input(input, { inputType: 'insertCompositionText', data: '漢', isComposing: true });
  });
  expect(client.sent.map((intent) => intent.kind)).toEqual(['compositionStart']);
  const start = client.sent[0];
  expect(start).toMatchObject({ kind: 'compositionStart', payload: {
    at: { source: 'confirmed', selection: selection() },
  } });
  if (start?.kind !== 'compositionStart') throw new Error('expected composition start');
  const id = start.payload.compositionId;
  expect(store.getSnapshot().composition).toEqual({ id, selection: selection(), text: '漢' });
  const selectionStart = input.selectionStart;
  const selectionEnd = input.selectionEnd;
  await act(async () => { client.ack(1, { frame: frame('abc', 1, 1) }); });
  expect(store.getSnapshot().confirmed!.frame.slideRevision).toBe(1);
  expect(store.getSnapshot().composition).toEqual({ id, selection: selection(), text: '漢' });
  expect(input.value).toBe('漢');
  expect(input.selectionStart).toBe(selectionStart);
  expect(input.selectionEnd).toBe(selectionEnd);
  expect(document.activeElement === input).toBe(true);
  expect(paintEcho.mock.calls.every(([view]) => view.echo === null)).toBe(true);
  await act(async () => {
    compose(input, 'compositionupdate', '漢字');
    input.value = '漢字';
    compose(input, 'compositionend', '漢字');
    beforeInput(input, 'insertText', '漢字');
    fireEvent.input(input, { inputType: 'insertText', data: '漢字' });
  });
  expect(client.sent).toHaveLength(2);
  expect(client.sent.filter((intent) => intent.kind === 'compositionCommit')).toEqual([
    { kind: 'compositionCommit', payload: expect.objectContaining({ compositionId: id, text: '漢字',
      at: { source: 'confirmed', selection: selection() } }) },
  ]);
  expect(store.getSnapshot().waitingInputCount).toBe(1);
  await act(async () => { client.ack(2, { selection: selection(3), frame: frame('a漢字bc', 2, 2) }); });
  expect(client.sent.filter((intent) => intent.kind === 'compositionEnd')).toHaveLength(1);
  expect(client.sent.find((intent) => intent.kind === 'compositionEnd')).toMatchObject({
    payload: { compositionId: id, at: { source: 'afterSeq', seq: 2 } },
  });
  expect(store.getSnapshot().composition).toBeNull();
  expect(store.getSnapshot().echo).toBeNull();
  await act(async () => {
    client.ack(3, { version: 'v2', selection: selection(3), frame: frame('a漢字bc', 3, 3),
      editState: { ...client.state.editState, canUndo: true } });
  });
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().presentedThroughSeq).toBe(3);
  expectNoMainModel(client);
});

test('p2_empty_composition_cancels_without_history', async () => {
  const { client, store, restoreLine } = setup();
  const view = render(<PeerlessTextInput store={store} />);
  const input = textarea(view.container);
  await act(async () => {
    compose(input, 'compositionstart', '');
    compose(input, 'compositionupdate', '仮');
    input.value = '仮';
    compose(input, 'compositionend', '');
    beforeInput(input, 'insertFromComposition', '');
    fireEvent.input(input, { inputType: 'insertFromComposition', data: '' });
  });
  expect(client.sent.map((intent) => intent.kind)).toEqual(['compositionStart', 'compositionEnd']);
  expect(client.sent[1]).toMatchObject({ kind: 'compositionEnd', payload: { cancelled: true } });
  expect(client.sent.some((intent) => intent.kind === 'compositionCommit')).toBe(false);
  expect(store.getSnapshot()).toMatchObject({ composition: null, echo: null,
    editState: { canUndo: false, canRedo: false } });
  expect(restoreLine).toHaveBeenCalled();
  await act(async () => { client.ack(2); });
  expect(store.getSnapshot().confirmed!.displayList).toEqual(JSON.parse(new TextDecoder().decode(frame().displayList)));
  expect(store.getSnapshot().editState.canUndo).toBe(false);
  expect(client.read).not.toHaveBeenCalled();
});

test('p2_composition_then_type_references_commit_and_keeps_character', async () => {
  const { client, store } = setup();
  const view = render(<PeerlessTextInput store={store} />);
  const input = textarea(view.container);
  await act(async () => {
    input.focus();
    compose(input, 'compositionstart', '');
    compose(input, 'compositionupdate', '漢');
    compose(input, 'compositionend', '漢');
    beforeInput(input, 'insertText', '!');
  });
  expect(client.sent.map((intent) => intent.kind)).toEqual([
    'compositionStart', 'compositionCommit',
  ]);
  expect(client.envelopes[1]).toMatchObject({ afterSeq: 1, payload: { at: { source: 'afterSeq', seq: 1 } } });
  expect(store.getSnapshot().waitingInputCount).toBe(2);
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().selection).toMatchObject({ anchor: 1, focus: 1 });
  await act(async () => { client.ack(2, { selection: selection(2), frame: frame('a漢bc', 1, 2) }); });
  expect(client.envelopes[2].afterSeq).toBe(2);
  expect(client.envelopes[3]).toMatchObject({ afterSeq: 3, payload: {
    text: '!', at: { source: 'afterSeq', seq: 3 },
  } });
  await act(async () => {
    client.ack(4, { version: 'v2', selection: selection(3), frame: frame('a漢!bc', 2, 4) });
  });
  expect(store.getSnapshot().pending).toHaveLength(0);
  const box = store.getSnapshot().confirmed!.displayList.primitives[0];
  if (box.kind !== 'textBox') throw new Error('Expected text box');
  expect(box.lines[0].runs[0].text).toBe('a漢!bc');
  expectNoMainModel(client);
});

test('p2_plain_paste_is_one_atomic_command', async () => {
  const { client, store } = setup();
  store.setSelection(selection(0, 2));
  const view = render(<PeerlessTextInput store={store} />);
  const input = textarea(view.container);
  const plain = 'A\n漢🙂';
  const getData = mock((type: string) => type === 'text/plain' ? plain : '<b>HTML</b>');
  let paste!: Event;
  await act(async () => {
    paste = beforeInput(input, 'insertFromPaste', null, {
      dataTransfer: { types: ['text/plain', 'text/html'], getData },
    });
    fireEvent.input(input, { inputType: 'insertFromPaste', data: plain });
  });
  expect(paste.defaultPrevented).toBe(true);
  expect(getData).toHaveBeenCalledWith('text/plain');
  expect(getData).not.toHaveBeenCalledWith('text/html');
  expect(client.sent).toHaveLength(1);
  expect(typeof client.sent[0]?.payload.inputTime).toBe('number');
  expect(client.sent[0]?.payload.undoGroup).toMatch(/^paste-\d+$/);
  expect(client.sent[0]).toMatchObject({ kind: 'replaceText', payload: {
    text: plain, at: { source: 'confirmed', selection: selection(0, 2) },
  } });
  expect(client.sent.some((intent) => intent.kind === 'deleteBackward' || intent.kind === 'deleteForward')).toBe(false);
  expect(store.getSnapshot().pending).toHaveLength(1);
  await act(async () => {
    client.ack(1, { version: 'v1', selection: selection(plain.length),
      frame: frame(`${plain}c`, 1, 1), editState: { ...client.state.editState, canUndo: true } });
  });
  expect(store.getSnapshot().selection).toMatchObject({ anchor: plain.length, focus: plain.length });
  expect(store.getSnapshot().editState.canUndo).toBe(true);
  expectNoMainModel(client);
});

function EditorShell({ store, editorRender, stripRender }: {
  store: PeerlessViewStore;
  editorRender: () => void;
  stripRender: () => void;
}) {
  useSyncExternalStore(store.subscribeStructure, () => store.getSnapshot().confirmed!.frame.slideId);
  editorRender();
  return <div><PeerlessTextInput store={store} /><ThumbnailStrip store={store} renderCount={stripRender} /></div>;
}

function ThumbnailStrip({ store, renderCount }: { store: PeerlessViewStore; renderCount: () => void }) {
  useSyncExternalStore(store.subscribeStructure, () => store.getSnapshot().thumbnails);
  renderCount();
  return <aside aria-label="Slides" />;
}

test('p2_bare_ack_does_not_render_editor_or_strip', async () => {
  const { client, store, paintFrame, paintEcho } = setup();
  const editorRender = mock(() => {});
  const stripRender = mock(() => {});
  const structural = mock(() => {});
  const toolbar = mock(() => {});
  store.subscribeStructure(structural);
  store.subscribeEditState(toolbar);
  const view = render(<EditorShell store={store} editorRender={editorRender} stripRender={stripRender} />);
  const input = textarea(view.container);
  await act(async () => { beforeInput(input, 'insertText', 'b'); });
  expect(client.sent).toHaveLength(1);
  expect(store.getSnapshot().echo).toBeNull();
  expect(store.getSnapshot().pending[0]?.echoed).toBe(false);
  expect(paintEcho.mock.calls.every(([view]) => view.echo === null)).toBe(true);
  const editors = editorRender.mock.calls.length;
  const strips = stripRender.mock.calls.length;
  const frames = paintFrame.mock.calls.length;
  const echoes = paintEcho.mock.calls.length;
  await act(async () => {
    client.ack(1, { version: 'v1', selection: selection(2), dirtySlideIds: ['slide'],
      editState: { ...client.state.editState, canUndo: true } });
  });
  expect(store.getSnapshot().appliedThroughSeq).toBe(1);
  expect(store.getSnapshot().presentedThroughSeq).toBe(0);
  expect(store.getSnapshot().echo).toBeNull();
  expect(editorRender).toHaveBeenCalledTimes(editors);
  expect(stripRender).toHaveBeenCalledTimes(strips);
  expect(structural).not.toHaveBeenCalled();
  expect(toolbar).toHaveBeenCalledTimes(1);
  expect(store.getSnapshot().editState.canUndo).toBe(true);
  expect(paintFrame).toHaveBeenCalledTimes(frames);
  expect(paintEcho).toHaveBeenCalledTimes(echoes);
  expect(client.thumbnail).not.toHaveBeenCalled();
  act(() => {
    client.deliverFrame(frame('abbc', 1, 1));
    expect(paintFrame).toHaveBeenCalledTimes(frames + 1);
    expect(store.getSnapshot().echo).toBeNull();
  });
  expect(editorRender).toHaveBeenCalledTimes(editors);
  expect(stripRender).toHaveBeenCalledTimes(strips);
  expectNoMainModel(client);
});

test('p3_arrow_keys_are_inert_after_terminal_failure', async () => {
  const { client, store } = setup();
  const view = render(<PeerlessTextInput store={store} />);
  const input = textarea(view.container);
  const setSelection = spyOn(store, 'setSelection').mockImplementation(() => {
    throw new Error('failed selection guard');
  });
  disposers.push(() => { setSelection.mockRestore(); });
  await act(async () => {
    for (const listener of client.failureListeners) listener(new CommandOutcomeUnknown(1, 'unknown outcome'));
  });
  const previous = store.getSnapshot().selection;
  await act(async () => {
    for (const key of ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End']) {
      expect(() => { fireEvent.keyDown(input, { key }); }).not.toThrow();
    }
  });
  expect(setSelection).not.toHaveBeenCalled();
  expect(client.sent).toEqual([]);
  expect(store.getSnapshot().selection).toBe(previous);
});

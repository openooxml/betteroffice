import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { StrictMode } from 'react';
import * as pptx from '@betteroffice/pptx';
import * as editor from '@betteroffice/pptx/editor';
import * as browserImage from '../../../pptx/src/render/browserImage';
import { createSessionHost } from '../../../../shared/office-session/host';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import { initWasm } from '../../../pptx/src/wasm/loader';
import { SessionFailure } from '../../../../shared/office-session/types';
import type {
  PeerlessSession, PresentationCommandAck, PresentationCommandClient, PresentationCommandFrame,
  PresentationCommandIntent, PresentationCommandOpenResult, PresentationCommandReadResult,
  PresentationCommandEnvelope, PresentationCommandFailure,
  PresentationCommandEvents, PresentationCommandMethods,
  PresentationCommandSelected, PresentationCommandSelection, PresentationCommandThumbnailResult, PptxReadRequest,
  SlideDisplayList, TextBoxPrimitive,
} from '@betteroffice/pptx';
import { PEERLESS_FLAG, PptxEditor, pptxPeerlessEditorLoader } from '../PptxEditor';
import type { PptxAsyncEditorApi } from './createPeerlessEditorApi';
import type { PeerlessConfirmedFrame, PeerlessViewSnapshot } from './peerlessViewStore';
import { peerlessPainterFactories, peerlessSessionOpener, type PeerlessPresentation } from './usePeerlessPresentation';
import { PeerlessSlideStrip } from './PeerlessSlideStrip';
import { commandFrameRaster, createCommandFrameImages, createCommandFramePainter,
  retainCommandFrameRaster } from './commandFramePainter';
import { isCommandAck, isCommandEnvelope } from '../../../../shared/office-session/commands/types';
import type { HostMessage } from '../../../../shared/office-session/protocol';
import { createPresentationSessionHost } from '../../../pptx/src/session/host';
import { createLocalEchoPainter } from './localEchoPainter';
import * as peerlessViewStore from './peerlessViewStore';

const { createPeerlessViewStore, commitComposition, PEERLESS_PENDING_INTENTS } = peerlessViewStore;
let echoEnabled = false;
mock.module('./peerlessViewStore', () => ({
  ...peerlessViewStore,
  createPeerlessViewStore: (...args: Parameters<typeof createPeerlessViewStore>) =>
    createPeerlessViewStore(args[0], args[1], echoEnabled ? { ...args[2], echoEnabled: true } : args[2]),
}));

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const file = new Uint8Array([1]);
let paints: PeerlessConfirmedFrame[];
let paintedViews: PeerlessViewSnapshot[];
let suppliedMedia: PresentationCommandFrame['media'][];
let duringPaint: (() => void) | undefined;
let duringDrawImage: ((source: unknown) => void) | undefined;
let canvasCalls: Map<HTMLCanvasElement, CanvasCall[]>;

interface CanvasCall {
  kind: string;
  args: readonly unknown[];
  style?: { font: unknown; color: unknown; letterSpacing: unknown };
}

function calls(canvas: HTMLCanvasElement): CanvasCall[] {
  let entries = canvasCalls.get(canvas);
  if (!entries) { entries = []; canvasCalls.set(canvas, entries); }
  return entries;
}

function kinds(entries: readonly CanvasCall[]): string[] {
  return entries.map((entry) => entry.kind);
}

function recordCanvases(): void {
  canvasCalls = new Map();
  const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
  const getContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (
    this: HTMLCanvasElement, type: string,
  ) {
    if (type !== '2d') return null;
    let context = contexts.get(this);
    if (!context) {
      const canvas = this;
      const properties: Record<PropertyKey, unknown> = { canvas,
        measureText: (text: string) => ({ width: text.length * 10 }),
        createLinearGradient: () => ({ addColorStop() {} }),
        createRadialGradient: () => ({ addColorStop() {} }),
      };
      context = new Proxy(properties, {
        get(target, key) {
          return key in target ? target[key] : (...args: unknown[]) => { calls(canvas).push({ kind: String(key), args,
            ...(key === 'fillText' ? { style: { font: target.font, color: target.fillStyle,
              letterSpacing: target.letterSpacing } } : {}) });
            if (key === 'drawImage') duringDrawImage?.(args[0]); };
        },
        set(target, key, value: unknown) { target[key] = value; return true; },
      }) as unknown as CanvasRenderingContext2D;
      contexts.set(canvas, context);
    }
    return context;
  } as typeof HTMLCanvasElement.prototype.getContext);
  restorers.push(() => getContext.mockRestore());
  for (const dimension of ['width', 'height'] as const) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, dimension);
    if (!descriptor?.get || !descriptor.set) throw new Error(`Expected canvas ${dimension} accessor`);
    const set = descriptor.set;
    Object.defineProperty(HTMLCanvasElement.prototype, dimension, { ...descriptor,
      set(this: HTMLCanvasElement, value: number) {
        calls(this).push({ kind: 'resize', args: [dimension, value] });
        set.call(this, value);
      },
    });
    restorers.push(() => Object.defineProperty(HTMLCanvasElement.prototype, dimension, descriptor));
  }
}

function expectRestored(canvas: HTMLCanvasElement, authoritative: HTMLCanvasElement): void {
  const entries = calls(canvas);
  expect(entries.some((entry) => entry.kind === 'drawImage' && entry.args[0] === authoritative)).toBe(true);
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index];
    if (entry.kind !== 'clearRect' && entry.kind !== 'resize') continue;
    expect(entries.slice(index + 1).some((next) => next.kind === 'drawImage' && next.args[0] === authoritative)).toBe(true);
  }
}

function expectIntact(canvas: HTMLCanvasElement, authoritative: HTMLCanvasElement, since: number): void {
  expectRestored(canvas, authoritative);
  const entries = calls(canvas);
  expect(kinds(entries.slice(since).filter((entry) => !['save', 'restore', 'setTransform'].includes(entry.kind)))).toEqual([]);
  const draws = entries.filter((entry) => entry.kind === 'drawImage');
  expect(draws[draws.length - 1].args[0] === authoritative).toBe(true);
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function textSelection(at: PresentationCommandSelected | null): PresentationCommandSelection {
  if (!at || at.kind === 'object') throw new Error('Expected text selection');
  return at;
}

const selection: PresentationCommandSelection = { slideId: 'slide', shapeId: 'box',
  storyId: 'story', anchor: 1, focus: 1, affinity: 'downstream' };

function initial(): PresentationCommandOpenResult & { frame: PresentationCommandFrame; activeSlideId: string } {
  const box: TextBoxPrimitive = {
    kind: 'textBox', objectId: 1, shapeId: 'box', storyId: 'story',
    x: 20, y: 40, w: 300, h: 40, anchor: 'top',
    paragraphs: [{ level: 0, runs: [{ text: 'abc', fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
    lines: [{ x: 20, y: 40, width: 30, height: 20, baseline: 55, start: 0, end: 3,
      caretStops: [0, 1, 2, 3].map((position) => ({ position, x: 20 + position * 10 })),
      runs: [{ text: 'abc', start: 0, end: 3, x: 20, width: 30, fontId: 0, fontFamily: 'Arial',
        fontSizePx: 16, bold: false, italic: false, underline: false, color: '#000000', glyphs: [] }] }],
  };
  const displayList: SlideDisplayList = { contractVersion: 1, width: 960, height: 540,
    background: { kind: 'solid', color: '#ffffff' }, primitives: [box] };
  const frame: PresentationCommandFrame = { slideId: 'slide', slideIndex: 0,
    slideRevision: 0, coversThroughSeq: 0, documentEpoch: 1, incarnation: 'worker', epoch: 1,
    version: 'v0', displayList: new TextEncoder().encode(JSON.stringify(displayList)).buffer, media: [] };
  return { session: 'session', incarnation: 'worker', epoch: 1, throughSeq: 0, version: 'v0',
    selection, editState: { readOnly: false, canUndo: false, canRedo: false,
      toolbar: { target: selection, storedMarks: {}, formatting: null, alignment: null,
        hyperlink: null, table: null, canEditText: true, ...{ echoRun: 'after' as const } } },
    outcomes: [], dirtySlideIds: [], frame, activeSlideId: 'slide',
    size: { width: 9144000, height: 5143500 },
    slides: [{ id: 'slide', index: 0, name: null, layoutPartPath: null }] };
}

class FakeCommandClient implements PresentationCommandClient {
  state = initial();
  text = 'abc';
  hold = false;
  refuse?: (intent: PresentationCommandEnvelope) => PresentationCommandAck['outcomes'][number]['reason'];
  readonly versions = new Map<string, number>([['v0', 0]]);
  readonly mutations: { seq: number; from: number; to: number; length: number }[] = [];
  readonly outcomes = new Map<number, PresentationCommandAck['outcomes'][number]>();
  readonly sent: PresentationCommandEnvelope[] = [];
  readonly replies = new Map<number, ReturnType<typeof deferred<PresentationCommandAck>>>();
  readonly selections = new Map<number, PresentationCommandSelection | null>();
  readonly listeners = new Set<(ack: PresentationCommandAck) => void>();
  readonly frames = new Set<(frame: PresentationCommandFrame) => void>();
  readonly failures = new Set<(error: PresentationCommandFailure) => void>();
  readonly media = new Map<string, ArrayBuffer>();
  readonly mediaListeners = new Set<(media: PresentationCommandFrame['media']) => void>();
  dispatch = mock((intent: PresentationCommandIntent, basis?: { epoch: number; version: string }): Promise<PresentationCommandAck> => {
    const seq = this.sent.length + 1;
    this.sent.push({ ...intent, session: this.state.session, incarnation: this.state.incarnation, seq,
      basis: basis ?? { epoch: this.state.epoch, version: this.state.version },
      ...('at' in intent.payload && intent.payload.at?.source === 'afterSeq' ? { afterSeq: intent.payload.at.seq } : {}) });
    if (this.hold) {
      const reply = deferred<PresentationCommandAck>();
      this.replies.set(seq, reply);
      return reply.promise;
    }
    return Promise.resolve(this.ack(seq));
  });
  ack(seq: number, frame?: PresentationCommandFrame): PresentationCommandAck {
    const { frame: _frame, ...base } = this.state;
    let selected = base.selection?.kind === 'object' ? null : base.selection;
    let readOnly = base.editState.readOnly;
    let version = base.version;
    const outcomes: PresentationCommandAck['outcomes'] = [];
    for (let sequence = base.throughSeq + 1; sequence <= seq; sequence += 1) {
      const intent = this.sent[sequence - 1];
      const basis = this.versions.get(intent.basis.version);
      if (basis === undefined || intent.basis.epoch !== base.epoch) throw new Error('Expected admitted basis');
      if (intent.afterSeq !== undefined) {
        const dependency = this.outcomes.get(intent.afterSeq);
        if (!dependency || dependency.status === 'refused' || dependency.status === 'cancelled') {
          const outcome: PresentationCommandAck['outcomes'][number] = {
            seq: sequence, status: 'cancelled', reason: 'dependency-refused', committed: false };
          outcomes.push(outcome);
          this.outcomes.set(sequence, outcome);
          continue;
        }
      }
      const refusal = this.refuse?.(intent);
      if (refusal) {
        const outcome: PresentationCommandAck['outcomes'][number] = { seq: sequence, status: 'refused', reason: refusal, committed: false };
        outcomes.push(outcome);
        this.outcomes.set(sequence, outcome);
        continue;
      }
      let committed = false;
      if (intent.kind === 'setReadOnly') readOnly = intent.payload.readOnly;
      else if (intent.kind === 'setSelection') selected = intent.payload.selection
        ? this.rebase(intent.payload.selection, basis) : null;
      else if (intent.kind === 'selectAtPoint') {
        const hit = pptx.hitGeometry(pptx.indexHitGeometry(pptx.decodeCommandDisplayList(this.state.frame)), intent.payload);
        if (hit?.kind === 'text') selected = { slideId: intent.payload.frame.slideId,
          shapeId: hit.shapeId, storyId: hit.storyId, anchor: intent.payload.extend && selected ? selected.anchor : hit.position, focus: hit.position,
          affinity: hit.affinity ?? 'downstream' };
      }
      else if (intent.kind === 'moveSelection') selected = pptx.navigateGeometry(
        pptx.indexGeometry(pptx.decodeCommandDisplayList(this.state.frame)), selected, intent.payload.motion, intent.payload.extend, this.text);
      else if (intent.kind === 'insertText' || intent.kind === 'replaceText' || intent.kind === 'compositionCommit') {
        const reference = intent.payload.at;
        const original = reference.source === 'confirmed' ? reference.selection : this.selections.get(reference.seq);
        const at = original && this.rebase(original, reference.source === 'confirmed' ? basis : reference.seq);
        if (!at) throw new Error('Expected mutation selection result');
        const from = Math.min(at.anchor, at.focus);
        const to = Math.max(at.anchor, at.focus);
        if (from < 0 || to > this.text.length) throw new Error('Expected valid text offsets');
        this.text = this.text.slice(0, from) + intent.payload.text + this.text.slice(to);
        const focus = Math.min(at.anchor, at.focus) + intent.payload.text.length;
        selected = { ...at, anchor: focus, focus };
        this.mutations.push({ seq: sequence, from, to, length: intent.payload.text.length });
        version = `v${sequence}`;
        this.versions.set(version, sequence);
        committed = true;
      }
      this.selections.set(sequence, selected);
      const outcome: PresentationCommandAck['outcomes'][number] = { seq: sequence, status: 'applied', committed };
      outcomes.push(outcome);
      this.outcomes.set(sequence, outcome);
    }
    if (!frame && outcomes.some((outcome) => this.sent[outcome.seq - 1]?.kind === 'selectAtPoint')) frame = this.state.frame;
    if (frame) frame = { ...coveringFrame(this, this.text, seq, frame), version };
    const ack: PresentationCommandAck = { ...base, version, throughSeq: seq, selection: selected,
      outcomes, dirtySlideIds: ['slide'], editState: { ...base.editState, readOnly }, ...(frame ? { frame } : {}) };
    this.state = { ...this.state, ...ack, activeSlideId: ack.frame?.slideId ?? this.state.activeSlideId, frame: ack.frame ?? this.state.frame };
    if (frame?.media.length) {
      for (const asset of frame.media) this.media.set(asset.assetId, asset.bytes);
      for (const listener of this.mediaListeners) listener(frame.media);
    }
    for (const listener of this.listeners) listener(ack);
    if (frame) for (const listener of this.frames) listener(frame);
    for (const [sequence, reply] of this.replies) {
      if (sequence <= seq) { reply.resolve(ack); this.replies.delete(sequence); }
    }
    return ack;
  }
  rebase(selection: PresentationCommandSelection, basis: number): PresentationCommandSelection {
    let at = selection;
    for (const mutation of this.mutations) {
      if (mutation.seq <= basis) continue;
      const map = (position: number) => position < mutation.from ? position
        : position > mutation.to ? position + mutation.length - (mutation.to - mutation.from)
        : mutation.from + mutation.length;
      at = { ...at, anchor: map(at.anchor), focus: map(at.focus) };
    }
    return at;
  }
  flush = mock((_seq?: number): Promise<PresentationCommandAck> => Promise.resolve(this.state));
  read = mock((_request?: PptxReadRequest): Promise<PresentationCommandReadResult> => Promise.resolve({
    ...this.state, result: { ok: true, version: this.state.version, slides: [], stories: [] },
  }));
  save = mock((): Promise<Uint8Array> => Promise.resolve(new TextEncoder().encode(this.text)));
  thumbnail = mock((slideId: string, knownRevision?: number): Promise<PresentationCommandThumbnailResult> =>
    Promise.resolve({ session: this.state.session, incarnation: this.state.incarnation,
      epoch: this.state.epoch, slideId, slideRevision: knownRevision ?? 0, status: 'unchanged' }));
  dispose = mock(() => Promise.resolve());
  subscribe(listener: (ack: PresentationCommandAck) => void) {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }
  subscribeFrame(listener: (frame: PresentationCommandFrame) => void) {
    this.frames.add(listener);
    return () => { this.frames.delete(listener); };
  }

  subscribeMedia(listener: (media: PresentationCommandFrame['media']) => void) {
    for (const asset of this.state.frame.media) this.media.set(asset.assetId, asset.bytes);
    this.mediaListeners.add(listener);
    listener([...this.media].map(([assetId, bytes]) => ({ assetId, bytes })));
    return () => { this.mediaListeners.delete(listener); };
  }
  subscribeFailure(listener: (error: PresentationCommandFailure) => void) {
    this.failures.add(listener);
    return () => { this.failures.delete(listener); };
  }
}

function session(held = false) {
  const client = new FakeCommandClient();
  const ready = deferred<PresentationCommandClient>();
  const listeners = new Set<(error: PresentationCommandFailure) => void>();
  let failure: PresentationCommandFailure | undefined;
  const owner: PeerlessSession = {
    ready: ready.promise,
    get failure() { return failure; },
    subscribeFailure(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    dispose: mock(() => client.dispose()),
  };
  if (!held) ready.resolve(client);
  const opener = spyOn(peerlessSessionOpener, 'open').mockReturnValue(owner);
  restorers.push(() => opener.mockRestore());
  return { client, owner, opener, admit: () => ready.resolve(client),
    crash() {
      failure = new SessionFailure('trap', 'worker trap', 'worker diagnostics');
      for (const listener of listeners) listener(failure);
      return failure;
    } };
}

beforeEach(() => {
  if (typeof Worker !== 'function') {
    const original = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
    Object.defineProperty(globalThis, 'Worker', { configurable: true, value: class {} });
    restorers.push(() => {
      if (original) Object.defineProperty(globalThis, 'Worker', original);
      else Reflect.deleteProperty(globalThis, 'Worker');
    });
  }
  paints = [];
  paintedViews = [];
  suppliedMedia = [];
  duringPaint = undefined;
  duringDrawImage = undefined;
  recordCanvases();
  for (const [property, extent] of [['clientWidth', 1000], ['clientHeight', 580]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, property);
    Object.defineProperty(HTMLElement.prototype, property, { configurable: true,
      get(this: HTMLElement) { return this.getAttribute('role') === 'application' ? extent : 0; } });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, property, original);
      else Reflect.deleteProperty(HTMLElement.prototype, property);
    });
  }
  const frame = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options) => {
    const painter = createCommandFramePainter(canvas, options);
    return { ...painter,
      paintFrame(value: PeerlessConfirmedFrame, view: PeerlessViewSnapshot, media?: PresentationCommandFrame['media']) {
        duringPaint?.();
        painter.paintFrame(value, view, media);
        paints.push(value);
        paintedViews.push(view);
        suppliedMedia.push(media ?? []);
      },
    };
  });
  restorers.push(() => frame.mockRestore());
});
afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('peerless editor chunk failures', () => {
  beforeEach(() => {
    const errors = spyOn(console, 'error').mockImplementation(() => {});
    restorers.push(() => errors.mockRestore());
  });

  async function rejectChunk(reject: (reason: unknown) => void, failure: Error) {
    try { await act(async () => { reject(failure); }); }
    catch (error) { if (error !== failure) throw error; }
  }

  test('keeps host siblings mounted and shows the editor error after a rejected import', async () => {
    const chunk = deferred<Awaited<ReturnType<typeof pptxPeerlessEditorLoader.load>>>();
    const load = spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValue(chunk.promise);
    restorers.push(() => load.mockRestore());
    const view = render(<><span>Host sibling</span>
      <PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} />
    </>);
    await rejectChunk(chunk.reject, new TypeError('Chunk fetch failed'));
    expect(view.queryByText('Host sibling')).not.toBeNull();
    expect(view.getByTestId('pptx-error').getAttribute('role')).toBe('alert');
    expect(view.getByTestId('pptx-error').textContent).toContain('Failed to load');
  });

  test('reports a typed chunk load error once even when the host callback throws', async () => {
    const chunk = deferred<Awaited<ReturnType<typeof pptxPeerlessEditorLoader.load>>>();
    const load = spyOn(pptxPeerlessEditorLoader, 'load').mockReturnValue(chunk.promise);
    restorers.push(() => load.mockRestore());
    const onError = mock((_error: Error) => { throw new Error('Host callback failed'); });
    const props = { fonts: [], file, [PEERLESS_FLAG]: true as const, onError };
    const view = render(<StrictMode><span>Host sibling</span><PptxEditor {...props} /></StrictMode>);
    const failure = new TypeError('Chunk fetch failed');
    await rejectChunk(chunk.reject, failure);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0][0]).toBeInstanceOf(Error);
    expect(onError.mock.calls[0][0]).toMatchObject({
      name: 'PptxWorkerEditorChunkLoadError', code: 'chunk-load-failed', cause: failure,
    });
    view.rerender(<StrictMode><span>Host sibling</span><PptxEditor {...props} file={new Uint8Array([4])} /></StrictMode>);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(view.queryByText('Host sibling')).not.toBeNull();
  });

  test('retries a transient import failure on a later mount without opening a main-thread peer', async () => {
    const value = session();
    const main = spyOn(editor, 'initWasm').mockImplementation(() => { throw new Error('Main wasm initialized'); });
    const model = spyOn(editor, 'openPresentation').mockImplementation(() => { throw new Error('Main peer opened'); });
    const chunk = deferred<Awaited<ReturnType<typeof pptxPeerlessEditorLoader.load>>>();
    const originalLoad = pptxPeerlessEditorLoader.load;
    const load = spyOn(pptxPeerlessEditorLoader, 'load').mockImplementation(originalLoad).mockReturnValueOnce(chunk.promise);
    restorers.push(() => main.mockRestore(), () => model.mockRestore(), () => load.mockRestore());
    const onError = mock((_error: Error) => {});
    const props = { fonts: [], file, [PEERLESS_FLAG]: true as const, onError };
    const failed = render(<PptxEditor {...props} />);
    await rejectChunk(chunk.reject, new TypeError('Transient chunk fetch failure'));
    failed.unmount();
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    render(<PptxEditor {...props} onReady={ready} />);
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(load).toHaveBeenCalledTimes(2);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(value.client.state.frame).not.toBeNull();
    expect(main).not.toHaveBeenCalled();
    expect(model).not.toHaveBeenCalled();
  });
});

test('p2_peerless_mount_ready_after_first_authoritative_paint', async () => {
  const value = session(true);
  const ready = mock((_api: PptxAsyncEditorApi) => { expect(paints.length).toBeGreaterThan(0); });
  duringPaint = () => { if (paints.length === 0) expect(ready).not.toHaveBeenCalled(); };
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  expect(ready).not.toHaveBeenCalled();
  expect(paints).toEqual([]);
  await act(async () => { value.admit(); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  duringPaint = undefined;
  const api = ready.mock.calls[0][0];
  await expect(api.whenHydrated()).resolves.toBeUndefined();
  expect(paints[0].frame).toBe(value.client.state.frame);
  expect(view.container.querySelector('textarea') === null).toBe(false);
  expect(value.client.read).not.toHaveBeenCalled();
  expect(value.client.thumbnail).not.toHaveBeenCalled();
  await expect(api.saveAsync()).resolves.toEqual(new TextEncoder().encode('abc'));
  await expect(api.readContent()).resolves.toMatchObject({ ok: true });
  const legacy = api as unknown as { handle: unknown; save(): unknown };
  expect(() => legacy.handle).toThrow(pptx.PptxPeerlessUnsupportedError);
  expect(() => legacy.save()).toThrow(pptx.PptxPeerlessUnsupportedError);
  view.unmount();
  expect(value.owner.dispose).toHaveBeenCalledTimes(1);
});

test('p2_peerless_mount_never_initializes_main_wasm', async () => {
  const value = session();
  const forbidden = () => { throw new Error('main wasm initialized'); };
  const main = spyOn(editor, 'initWasm').mockImplementation(forbidden);
  const model = spyOn(editor, 'openPresentation').mockImplementation(forbidden);
  const instantiate = spyOn(WebAssembly, 'instantiate').mockImplementation(forbidden);
  const instantiateStreaming = spyOn(WebAssembly, 'instantiateStreaming').mockImplementation(forbidden);
  const compile = spyOn(WebAssembly, 'compile').mockImplementation(forbidden);
  const compileStreaming = spyOn(WebAssembly, 'compileStreaming').mockImplementation(forbidden);
  restorers.push(() => main.mockRestore(), () => model.mockRestore(),
    () => instantiate.mockRestore(), () => instantiateStreaming.mockRestore(),
    () => compile.mockRestore(), () => compileStreaming.mockRestore());
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const input = view.container.querySelector('textarea');
  if (!input) throw new Error('Expected peer-less text input');
  await act(async () => {
    input.focus();
    const event = new Event('beforeinput', { bubbles: true, cancelable: true });
    Object.defineProperties(event, { inputType: { value: 'insertText' }, data: { value: 'x' } });
    fireEvent(input, event);
  });
  expect(value.client.dispatch).toHaveBeenCalledTimes(1);
  expect(value.client.dispatch.mock.calls[0][0]).toMatchObject({ kind: 'insertText', payload: { text: 'x' } });
  expect(value.client.read).not.toHaveBeenCalled();
  expect(main).not.toHaveBeenCalled();
  expect(model).not.toHaveBeenCalled();
  expect(instantiate).not.toHaveBeenCalled();
  expect(instantiateStreaming).not.toHaveBeenCalled();
  expect(compile).not.toHaveBeenCalled();
  expect(compileStreaming).not.toHaveBeenCalled();
});

test('p2_refused_api_proposal_rejects_to_caller_without_onError', async () => {
  const value = session();
  value.client.refuse = (intent) => intent.kind === 'invoke' && intent.payload.method === 'propose' ? 'invalid-payload' : undefined;
  const errors = mock((_error: Error) => {});
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onError={errors} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const access = await ready.mock.calls[0][0].handleAsync();
  const edit = { type: 'insertSlide', index: 1 } as unknown as editor.ProposalEdit;
  await act(async () => {
    await expect(access.propose('Writer', null, [edit])).rejects.toMatchObject({ code: 'invalid-arguments' });
  });
  expect(errors).not.toHaveBeenCalled();
  expect(view.getByRole('application').getAttribute('aria-readonly')).toBe('false');
  expect(view.queryByRole('alert')).toBe(null);
  await act(async () => { await access.setSlideNotes('slide', 'notes'); });
  expect(value.client.outcomes.get(value.client.sent.length)).toMatchObject({ status: 'applied' });
  let failure!: PresentationCommandFailure;
  await act(async () => { failure = value.crash(); });
  expect(errors).toHaveBeenCalledTimes(1);
  expect(errors.mock.calls[0][0]).toMatchObject({ code: 'editor-failed', cause: failure });
});

test('p2_worker_failure_goes_read_only_with_typed_onError', async () => {
  const value = session();
  value.client.hold = true;
  const errors = mock((_error: Error) => {});
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onError={errors} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const lastFrame = paints[paints.length - 1];
  const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
  const raster = commandFrameRaster(lastFrame);
  if (!raster) throw new Error('Expected authoritative raster');
  const input = view.container.querySelector('textarea');
  if (!input) throw new Error('Expected peer-less text input');
  const intactSince = calls(canvas).length;
  await act(async () => {
    input.focus();
    const event = new Event('beforeinput', { bubbles: true, cancelable: true });
    Object.defineProperties(event, { inputType: { value: 'insertText' }, data: { value: 'X' } });
    fireEvent(input, event);
  });
  expect(kinds(calls(canvas).filter((entry) => entry.kind === 'fillText'))).toEqual([]);
  const authoritativeCalls = calls(raster.canvas).length;
  let failure!: PresentationCommandFailure;
  await act(async () => { failure = value.crash(); });
  expect(errors).toHaveBeenCalledTimes(1);
  expect(errors.mock.calls[0][0]).toBeInstanceOf(pptx.PptxWorkerEditorFailedError);
  expect(errors.mock.calls[0][0]).toMatchObject({ code: 'editor-failed', cause: failure });
  expect(failure).toMatchObject({ code: 'trap', diagnostics: 'worker diagnostics' });
  expect(view.getByRole('application').getAttribute('aria-readonly')).toBe('true');
  expect(view.getByTestId('pptx-slide-canvas') !== undefined).toBe(true);
  expect(paints[paints.length - 1].frame).toBe(lastFrame.frame);
  expectIntact(canvas, raster.canvas, intactSince);
  expect(calls(raster.canvas)).toHaveLength(authoritativeCalls);
  expect(calls(canvas).slice(intactSince).some((entry) => entry.kind === 'resize')).toBe(false);
  await act(async () => {
    const event = new Event('beforeinput', { bubbles: true, cancelable: true });
    Object.defineProperties(event, { inputType: { value: 'insertText' }, data: { value: 'x' } });
    fireEvent(input, event);
    fireEvent.compositionStart(input, { data: '' });
    fireEvent.compositionEnd(input, { data: 'x' });
  });
  expect(value.client.dispatch).toHaveBeenCalledTimes(1);
  await expect(ready.mock.calls[0][0].saveAsync()).rejects.toBe(errors.mock.calls[0][0]);
  expect(value.opener).toHaveBeenCalledTimes(1);
  expect(ready).toHaveBeenCalledTimes(1);
});

function coveringFrame(client: FakeCommandClient, text: string, seq: number,
  basis = client.state.frame): PresentationCommandFrame {
  const list = JSON.parse(new TextDecoder().decode(basis.displayList)) as SlideDisplayList;
  const primitives = list.primitives.map((primitive) => primitive.kind !== 'textBox' ? primitive : {
    ...primitive, lines: primitive.lines.map((line) => ({ ...line, end: text.length, width: text.length * 10,
      caretStops: Array.from({ length: text.length + 1 }, (_, position) => ({ position, x: line.x + position * 10 })),
      runs: line.runs.map((run) => ({ ...run, text, end: text.length, width: text.length * 10 })),
    })),
  });
  return { ...basis, coversThroughSeq: seq, slideRevision: client.state.frame.slideRevision + 1,
    displayList: new TextEncoder().encode(JSON.stringify({ ...list, primitives })).buffer };
}

function pictureFrame(frame: PresentationCommandFrame, assetId: string): PresentationCommandFrame {
  const list = JSON.parse(new TextDecoder().decode(frame.displayList)) as SlideDisplayList;
  list.primitives.push({ kind: 'image', objectId: 3, name: 'Picture', assetId, x: 400, y: 100, w: 100, h: 100 });
  return { ...frame, media: [{ assetId, bytes: new Uint8Array([1]).buffer }],
    displayList: new TextEncoder().encode(JSON.stringify(list)).buffer };
}

function heldImage() {
  const image = deferred<CanvasImageSource | null>();
  const resolve = Object.assign((_assetId: string) => image.promise, { release: mock(() => {}) });
  const cache = spyOn(editor, 'browserFrameImages').mockReturnValue({ resolve: () => resolve, dispose: mock(() => {}) });
  restorers.push(() => cache.mockRestore());
  return { image, release() {
    const source = new Image();
    source.width = source.height = 100;
    image.resolve(source);
    return source;
  } };
}

test.each(['painted', 'disposed', 'failed', 'stale', 'failed-layer'] as const)('p2_parallel_tiles_release_shared_images_after_all_draws (%s)', async (outcome) => {
  const client = new FakeCommandClient();
  const list = JSON.parse(new TextDecoder().decode(client.state.frame.displayList)) as SlideDisplayList;
  const closedBeforeDraw: boolean[] = [];
  const pictureDrawn = (source: CanvasImageSource) => [...canvasCalls.values()].some((entries) =>
    entries.some((entry) => entry.kind === 'drawImage' && entry.args[0] === source));
  const images = Array.from({ length: 26 }, (_, index) => {
    const pending = deferred<CanvasImageSource>();
    const source = { width: 100, height: 100, close: mock(() => {
      closedBeforeDraw.push(!pictureDrawn(source));
    }) };
    list.primitives.push({ kind: 'image', objectId: index + 2, name: `Picture ${index}`, assetId: `image:${index}`,
      x: 400, y: 100, w: 100, h: 100 });
    return { pending, source };
  });
  const drawCount = () => images.filter(({ source }) => pictureDrawn(source)).length;
  const decode = spyOn(browserImage, 'decodeBrowserPresentationImage').mockImplementation((bytes) =>
    images[bytes[0]].pending.promise);
  const frameImages = createCommandFrameImages();
  const pin = frameImages.pin;
  const releasedAfter: number[] = [];
  const release = mock(() => { releasedAfter.push(drawCount()); });
  frameImages.pin = (frame, supplied) => {
    const media = pin(frame, supplied);
    return { ...media, retain() {
      const held = media.retain();
      let released = false;
      const releaseHeld = () => {
        if (released) return;
        released = true;
        release(); held.release();
      };
      return { ...held, resolve: Object.assign((assetId: string) => held.resolve(assetId), { release: releaseHeld }),
        release: releaseHeld };
    } };
  };
  restorers.push(() => decode.mockRestore());
  client.state.frame = { ...client.state.frame,
    media: images.map((_, index) => ({ assetId: `image:${index}`, bytes: new Uint8Array([index]).buffer })),
    displayList: new TextEncoder().encode(JSON.stringify(list)).buffer };
  const canvas = document.createElement('canvas');
  let current = true;
  const painter = createCommandFramePainter(canvas, { images: frameImages, current: () => current });
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
  let consumer: ReturnType<typeof retainCommandFrameRaster>;
  try {
    const view = store.getSnapshot();
    if (outcome === 'failed') {
      duringDrawImage = (source) => {
        if (source === images[0].source) throw new Error('Paint failed');
      };
    }
    if (outcome === 'failed-layer') {
      const paint = pptx.paintSlide;
      const layer = spyOn(editor, 'paintSlide').mockImplementation(async (...args) => {
        await paint(...args);
        if (args[1].primitives.some((primitive) => primitive.kind === 'image' && primitive.assetId === 'image:0')) {
          throw new Error('Paint failed');
        }
      });
      restorers.push(() => layer.mockRestore());
    }
    painter.paintFrame(view.confirmed!, view);
    const painted = painter.whenPainted();
    const failure = painted.then(() => null, (error: unknown) => error as Error);
    const settled = painted.catch(() => false);
    if (outcome === 'disposed' || outcome === 'stale') {
      consumer = retainCommandFrameRaster(view.confirmed!, 1);
      expect(consumer).toBeDefined();
      if (outcome === 'disposed') { painter.dispose(); frameImages.dispose(); }
      else { current = false; painter.cancel(); }
    }
    await act(async () => {});
    expect(decode).toHaveBeenCalledTimes(images.length);
    expect(release).not.toHaveBeenCalled();
    await act(async () => {
      for (const image of images.slice(0, -1)) image.pending.resolve(image.source);
    });
    expect(drawCount()).toBe(images.length - 1);
    expect(release).not.toHaveBeenCalled();
    expect(images[0].source.close).not.toHaveBeenCalled();
    images[images.length - 1].pending.resolve(images[images.length - 1].source);
    await expect(settled).resolves.toBe(outcome === 'painted');
    if (outcome === 'failed' || outcome === 'failed-layer') expect((await failure)?.message).toBe('Paint failed');
    painter.freeze();
    expect(drawCount()).toBe(images.length);
    expect(release).toHaveBeenCalledTimes(1);
    expect(releasedAfter).toEqual([images.length]);
    expect(images[0].source.close).toHaveBeenCalledTimes(1);
    expect(closedBeforeDraw.every((closed) => !closed)).toBe(true);
  } finally {
    for (const image of images) image.pending.resolve(image.source);
    store.dispose();
    painter.dispose();
    consumer?.release();
    frameImages.dispose();
  }
});

test('p2_cancelled_raster_releases_decoded_image_while_other_decode_is_pending', async () => {
  const client = new FakeCommandClient();
  const list = JSON.parse(new TextDecoder().decode(client.state.frame.displayList)) as SlideDisplayList;
  const images = Array.from({ length: 2 }, (_, index) => {
    const pending = deferred<CanvasImageSource>();
    const source = { width: 100, height: 100, close: mock(() => {}) };
    list.primitives.push({ kind: 'image', objectId: index + 2, name: `Picture ${index}`, assetId: `image:${index}`,
      x: 400, y: 100, w: 100, h: 100 });
    return { pending, source };
  });
  const pictureDraws = () => [...canvasCalls.values()].flatMap((entries) => entries
    .filter((entry) => entry.kind === 'drawImage' && images.some(({ source }) => entry.args[0] === source))
    .map((entry) => entry.args[0]));
  const decode = spyOn(browserImage, 'decodeBrowserPresentationImage').mockImplementation((bytes) =>
    images[bytes[0]].pending.promise);
  restorers.push(() => decode.mockRestore());
  client.state.frame = { ...client.state.frame,
    media: images.map((_, index) => ({ assetId: `image:${index}`, bytes: new Uint8Array([index]).buffer })),
    displayList: new TextEncoder().encode(JSON.stringify(list)).buffer };
  const frameImages = createCommandFrameImages();
  const painter = createCommandFramePainter(document.createElement('canvas'), { images: frameImages });
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
  try {
    const view = store.getSnapshot();
    painter.paintFrame(view.confirmed!, view);
    let settled = false;
    const painted = painter.whenPainted().then((result) => { settled = true; return result; });
    await act(async () => { images[0].pending.resolve(images[0].source); });
    expect(decode).toHaveBeenCalledTimes(2);
    expect(pictureDraws()).toEqual([images[0].source]);
    expect(images.map(({ source }) => source.close.mock.calls.length)).toEqual([0, 0]);
    painter.cancel();
    frameImages.dispose();
    expect(images[0].source.close).toHaveBeenCalledTimes(1);
    expect(images[1].source.close).not.toHaveBeenCalled();
    await act(async () => {});
    expect(settled).toBe(false);
    await act(async () => { images[1].pending.resolve(images[1].source); await painted; });
    await expect(painted).resolves.toBe(false);
    expect(images.map(({ source }) => source.close.mock.calls.length)).toEqual([1, 1]);
    expect(pictureDraws()).toEqual([images[0].source]);
  } finally {
    for (const image of images) image.pending.resolve(image.source);
    store.dispose();
    painter.dispose();
    frameImages.dispose();
  }
});

test('p2_ready_and_thumbnails_wait_for_async_image_tiles', async () => {
  const value = session();
  value.client.state.frame = pictureFrame(value.client.state.frame, 'first');
  value.client.state.slides.push({ id: 'other', index: 1, name: null, layoutPartPath: null });
  const image = heldImage();
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(paints.length).toBeGreaterThan(0));
  await act(async () => { await Promise.resolve(); });
  expect(ready).not.toHaveBeenCalled();
  expect(value.client.thumbnail).not.toHaveBeenCalled();
  let source!: CanvasImageSource;
  await act(async () => { source = image.release(); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  expect([...canvasCalls.values()].some((entries) => entries.some((entry) =>
    entry.kind === 'drawImage' && entry.args[0] === source))).toBe(true);
  await ready.mock.calls[0][0].whenHydrated();
  await waitFor(() => expect(value.client.thumbnail).toHaveBeenCalledTimes(1));
});

test('p2_current_paint_gates_hydration_and_snapshot_without_delaying_frame_install', async () => {
  const value = session();
  value.client.hold = true;
  const image = heldImage();
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const changed = mock((_snapshot: pptx.DeckSnapshot) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} onChange={changed} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const api = ready.mock.calls[0][0];
  const input = view.container.querySelector('textarea');
  if (!input) throw new Error('Expected text input');
  let hydrated = false;
  let completion!: Promise<void>;
  await act(async () => {
    const event = new Event('beforeinput', { bubbles: true, cancelable: true });
    Object.defineProperties(event, { inputType: { value: 'insertText' }, data: { value: 'X' } });
    fireEvent(input, event);
    value.client.ack(1, pictureFrame(coveringFrame(value.client, 'aXbc', 1), 'late'));
    expect(paints[paints.length - 1]?.frame.coversThroughSeq).toBe(1);
    completion = api.whenHydrated().then(() => { hydrated = true; });
  });
  expect(hydrated).toBe(false);
  expect(value.client.read).not.toHaveBeenCalled();
  expect(changed).not.toHaveBeenCalled();
  await act(async () => { image.release(); await completion; });
  expect(hydrated).toBe(true);
  await waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
  expect(value.client.text).toBe('aXbc');
});

test('p2_authoritative_paint_survives_covering_frame_focus_and_selection_damage', async () => {
  const value = session();
  value.client.hold = true;
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
  const input = view.container.querySelector('textarea');
  if (!input) throw new Error('Expected peer-less text input');
  const initialRaster = commandFrameRaster(paints[0]);
  if (!initialRaster) throw new Error('Expected initial raster');
  expect(initialRaster.canvas === canvas).toBe(false);
  expectRestored(canvas, initialRaster.canvas);
  await act(async () => {
    input.focus();
    const event = new Event('beforeinput', { bubbles: true, cancelable: true });
    Object.defineProperties(event, { inputType: { value: 'insertText' }, data: { value: 'X' } });
    fireEvent(input, event);
  });
  await act(async () => { value.client.ack(1, coveringFrame(value.client, 'aXbc', 1)); });
  const frame = paints[paints.length - 1];
  if (!frame) throw new Error('Expected covering frame');
  const raster = commandFrameRaster(frame);
  if (!raster) throw new Error('Expected covering raster');
  expectRestored(canvas, raster.canvas);
  expect([...canvasCalls.values()].some((entries) => entries.some((entry) =>
    entry.kind === 'fillText' && entry.args[0] === 'aXbc'))).toBe(true);
  const authoritativeCalls = calls(raster.canvas).length;
  const intactSince = calls(canvas).length;
  await act(async () => { input.blur(); });
  expectIntact(canvas, raster.canvas, intactSince);
  expect(calls(canvas).slice(intactSince).some((entry) => entry.kind === 'resize')).toBe(false);
  calls(canvas).length = 0;
  await act(async () => {
    input.focus();
    fireEvent.keyDown(input, { key: 'ArrowRight', shiftKey: true });
  });
  expect(calls(canvas).some((entry) => entry.kind === 'fillRect'
    && entry.args.join(',') === '40,40,10,20')).toBe(true);
  calls(canvas).length = 0;
  await act(async () => {
    input.blur();
  });
  expectRestored(canvas, raster.canvas);
  expect(calls(raster.canvas)).toHaveLength(authoritativeCalls);
});

test('p2_save_flush_version_and_reads_wait_for_capped_input', async () => {
  const value = session();
  value.client.hold = true;
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const input = view.container.querySelector('textarea');
  if (!input) throw new Error('Expected peer-less text input');
  await act(async () => {
    for (let index = 0; index <= PEERLESS_PENDING_INTENTS; index += 1) {
      const event = new Event('beforeinput', { bubbles: true, cancelable: true });
      Object.defineProperties(event, { inputType: { value: 'insertText' }, data: { value: index === 256 ? 'LAST' : 'a' }, timeStamp: { value: 100 } });
      fireEvent(input, event);
    }
  });
  expect(value.client.sent).toHaveLength(1);
  const api = ready.mock.calls[0][0];
  const operations = [api.saveAsync(), api.flushPendingInput(), api.version(), api.readContent()];
  await act(async () => { await Promise.resolve(); });
  expect(value.client.save).not.toHaveBeenCalled();
  expect(value.client.flush).not.toHaveBeenCalled();
  expect(value.client.read).not.toHaveBeenCalled();
  await act(async () => { value.client.ack(1, coveringFrame(value.client, 'aabc', 1)); });
  expect(value.client.sent).toHaveLength(2);
  expect(value.client.sent[1].afterSeq).toBe(1);
  expect(value.client.sent[1]).toMatchObject({ basis: { version: 'v0' },
    payload: { text: `${'a'.repeat(255)}LAST`, at: { source: 'afterSeq', seq: 1 } } });
  expect(value.client.save).not.toHaveBeenCalled();
  expect(value.client.flush).not.toHaveBeenCalled();
  expect(value.client.read).not.toHaveBeenCalled();
  await act(async () => { value.client.ack(2, coveringFrame(value.client, `a${'a'.repeat(256)}LASTbc`, 2)); });
  const [saved] = await Promise.all(operations);
  if (!(saved instanceof Uint8Array)) throw new Error('Expected saved bytes');
  expect(new TextDecoder().decode(saved)).toBe(`a${'a'.repeat(256)}LASTbc`);
  expect(value.client.save).toHaveBeenCalledTimes(1);
  expect(value.client.flush).toHaveBeenCalledTimes(2);
  expect(value.client.read).toHaveBeenCalledTimes(1);
});

test('p2_read_only_false_true_false_enqueues_both_transitions', async () => {
  const value = session();
  value.client.hold = true;
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} readOnly={false} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  view.rerender(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} readOnly onReady={ready} />);
  expect(value.client.sent.map((intent) => intent.payload)).toMatchObject([{ readOnly: true }]);
  view.rerender(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} readOnly={false} onReady={ready} />);
  expect(value.client.sent.map((intent) => intent.payload)).toMatchObject([{ readOnly: true }, { readOnly: false }]);
  expect(value.client.sent.every((intent) => intent.afterSeq === undefined)).toBe(true);
  await act(async () => { value.client.ack(1); value.client.ack(2); });
  expect(value.client.state.editState.readOnly).toBe(false);
  expect(view.getByRole('application').getAttribute('aria-readonly')).toBe('false');
});

async function thumbnailMediaTyping() {
  const value = session();
  value.client.hold = true;
  value.client.state.slides.push({ id: 'other', index: 1, name: null, layoutPartPath: null });
  value.client.thumbnail.mockImplementation((slideId) => Promise.resolve({
    session: value.client.state.session, incarnation: value.client.state.incarnation,
    epoch: value.client.state.epoch, slideId, slideRevision: 1, status: 'frame',
    frame: { ...value.client.state.frame, slideId, slideIndex: 1, slideRevision: 1,
      media: [{ assetId: 'shared-image', bytes: new Uint8Array([1, 2]).buffer }] },
  }));
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(value.client.thumbnail).toHaveBeenCalledTimes(1));
  await waitFor(() => expect(paints.some((frame) => frame.frame.slideId === 'other')).toBe(true));
  const input = view.container.querySelector('textarea');
  if (!input) throw new Error('Expected peer-less text input');
  await act(async () => { fireEvent.keyDown(input, { key: 'ArrowRight' }); });
  await act(async () => { value.client.ack(1, coveringFrame(value.client, 'abc', 1)); });
  const index = paints.length - 1;
  expect(suppliedMedia.flat()).toContainEqual({ assetId: 'shared-image', bytes: new Uint8Array([1, 2]).buffer });
  expect(paints[index].frame.slideId).toBe('slide');
  expect(suppliedMedia[index]).toEqual([]);
  expect(paints[index]).toBe(paintedViews[index].confirmed!);
  expect(paintedViews[index]).toMatchObject({ presentedThroughSeq: 1, pending: [] });
  const raster = commandFrameRaster(paintedViews[index].confirmed!);
  if (!raster) throw new Error('Expected raster under store frame identity');
  const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
  calls(canvas).length = 0;
  await act(async () => {
    const event = new Event('beforeinput', { bubbles: true, cancelable: true });
    Object.defineProperties(event, { inputType: { value: 'insertText' }, data: { value: 'b' } });
    fireEvent(input, event);
  });
  expect(value.client.sent[1]).toMatchObject({ kind: 'insertText', payload: {
    at: { source: 'confirmed', selection: { anchor: 2, focus: 2 } } } });
  expect(value.client.sent[1].afterSeq).toBeUndefined();
  return { canvas, raster };
}

test('p2_thumbnail_media_preserves_confirmed_identity_and_echo_plate', async () => {
  echoEnabled = true;
  try {
    const { canvas, raster } = await thumbnailMediaTyping();
    expect(calls(canvas).filter((entry) => entry.kind === 'fillText').map((entry) => entry.args[0]).join('')).toBe('abbc');
    expect(calls(canvas).some((entry) => entry.kind === 'drawImage' && entry.args[0] !== raster.canvas)).toBe(true);
  } finally { echoEnabled = false; }
});

test('p2_thumbnail_media_preserves_confirmed_identity_without_echo_plate', async () => {
  const { canvas, raster } = await thumbnailMediaTyping();
  expect(kinds(calls(canvas).filter((entry) => entry.kind === 'fillText'))).toEqual([]);
  expect(calls(canvas).some((entry) => entry.kind === 'drawImage' && entry.args[0] !== raster.canvas)).toBe(false);
});

test('p2_shape_tile_includes_thick_stroke_and_line_end_extents', async () => {
  const value = session();
  const list: SlideDisplayList = { contractVersion: 1, width: 960, height: 540, primitives: [{
    kind: 'shape', objectId: 2, shapeId: 'thick', name: 'Thick line', geometry: 'line',
    x: 100, y: 100, w: 100, h: 50,
    path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 1 }],
    stroke: { color: '#000000', width: 20 },
  }, {
    kind: 'shape', objectId: 3, shapeId: 'marker', name: 'Marker line', geometry: 'line',
    x: 400, y: 100, w: 100, h: 50,
    path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 1 }],
    stroke: { color: '#000000', width: 2, tailEnd: { kind: 'diamond', width: 40, length: 60 } },
  }] };
  value.client.state.frame = { ...value.client.state.frame,
    displayList: new TextEncoder().encode(JSON.stringify(list)).buffer };
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const raster = commandFrameRaster(paints[0]);
  if (!raster) throw new Error('Expected shape raster');
  const draws = calls(raster.canvas).filter((entry) => entry.kind === 'drawImage');
  expect(draws).toHaveLength(2);
  for (const [index, draw] of draws.entries()) {
    if (!(draw.args[0] instanceof HTMLCanvasElement)) throw new Error('Expected shape tile');
    const tile = draw.args[0];
    const x = Number(draw.args[1]); const y = Number(draw.args[2]);
    const left = index === 0 ? 100 : 400;
    const spread = index === 0 ? 10 : 60;
    expect(x).toBeLessThanOrEqual(left - spread);
    expect(y).toBeLessThanOrEqual(100 - (index === 0 ? 10 : 20));
    expect(x + tile.width).toBeGreaterThanOrEqual(left + 100 + spread);
    expect(y + tile.height).toBeGreaterThanOrEqual(150 + (index === 0 ? 10 : 20));
    expect(calls(tile).some((entry) => entry.kind === 'stroke')).toBe(true);
  }
});

async function deck(text: string): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  const paragraphs = text.split('\n').map((paragraph) => `<a:p><a:r><a:rPr sz="1200"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${paragraph}</a:t></a:r></a:p>`).join('');
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="8000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="none"/><a:lstStyle/>${paragraphs}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function expandedDeck(text: string, slideCount: number, fields: boolean): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await deck(text));
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const source = await zip.file('ppt/slides/slide1.xml')!.async('string');
  if (fields) {
    const shape = source.match(/<p:sp>[\s\S]*?<\/p:sp>/)?.[0];
    if (!shape) throw new Error('Expected text shape');
    const plain = shape.replace('id="2"', 'id="3"').replace('y="100000"', 'y="2000000"')
      .replace(/<a:t>[\s\S]*?<\/a:t>/, '<a:t>abc</a:t>');
    const field = source.replace(/<a:r>[\s\S]*?<\/a:r>/,
      '<a:r><a:t>A</a:t></a:r><a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum"><a:t>12</a:t></a:fld><a:r><a:t>B</a:t></a:r>');
    zip.file('ppt/slides/slide1.xml', field.replace('</p:spTree>', `${plain}</p:spTree>`));
  }
  if (slideCount > 1) {
    const types = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', types.replace('</Types>', Array.from({ length: slideCount - 1 }, (_, index) =>
      `<Override PartName="/ppt/slides/slide${index + 2}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') + '</Types>'));
    const presentation = await zip.file('ppt/presentation.xml')!.async('string');
    zip.file('ppt/presentation.xml', presentation.replace('</p:sldIdLst>', Array.from({ length: slideCount - 1 }, (_, index) =>
      `<p:sldId id="${257 + index}" r:id="rId${index + 2}"/>`).join('') + '</p:sldIdLst>'));
    zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${Array.from({ length: slideCount }, (_, index) =>
      `<Relationship Id="rId${index + 1}" Type="${rel}/slide" Target="slides/slide${index + 1}.xml"/>`).join('')}</Relationships>`);
    for (let index = 2; index <= slideCount; index += 1) zip.file(`ppt/slides/slide${index}.xml`, source);
  }
  return zip.generateAsync({ type: 'uint8array' });
}

async function echoStyleDeck(paragraph: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await deck('alpha'));
  const path = 'ppt/slides/slide1.xml';
  const source = await zip.file(path)!.async('string');
  zip.file(path, source.replace(/<a:p>.*?<\/a:p>/, paragraph));
  return zip.generateAsync({ type: 'uint8array' });
}

function styledEchoParagraph(size = 42): string {
  return `<a:p><a:pPr><a:defRPr b="1" sz="1800"><a:solidFill><a:srgbClr val="111827"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr b="0" sz="${size * 100}"><a:solidFill><a:srgbClr val="3478CC"/></a:solidFill><a:latin typeface="Liberation Sans"/></a:rPr><a:t>alpha</a:t></a:r></a:p>`;
}

function echoView(client: PresentationCommandClient) {
  const canvas = document.createElement('canvas');
  const framePainter = createCommandFramePainter(canvas);
  const painter = createLocalEchoPainter(canvas);
  const store = createPeerlessViewStore(client, { ...framePainter, ...painter });
  expect(commandFrameRaster(store.getSnapshot().confirmed!) !== undefined).toBe(true);
  return { canvas, painter, store, dispose() { painter.dispose(); framePainter.dispose(); store.dispose(); } };
}

function frameRuns(frame: PresentationCommandFrame, storyId: string) {
  return pptx.decodeCommandDisplayList(frame).primitives.flatMap((primitive) =>
    primitive.kind === 'textBox' && primitive.storyId === storyId ? primitive.lines.flatMap((line) => line.runs) : []);
}

function expectNoEcho(canvas: HTMLCanvasElement, view: PeerlessViewSnapshot): void {
  expect(view.echo).toBeNull();
  expect(kinds(calls(canvas).filter((call) => call.kind === 'fillText'))).toEqual([]);
}

function expectEchoStyle(
  canvas: HTMLCanvasElement, view: PeerlessViewSnapshot, frame: PresentationCommandFrame,
  start: number, text: string,
): void {
  const list = pptx.decodeCommandDisplayList(frame);
  const box = list.primitives.find((primitive) => primitive.kind === 'textBox'
    && primitive.storyId === view.selection?.storyId);
  expect(box?.kind).toBe('textBox');
  if (!view.echo) {
    expectNoEcho(canvas, view);
    return;
  }
  expect(calls(canvas).filter((call) => call.kind === 'fillText').length).toBeGreaterThan(0);
  if (box?.kind !== 'textBox' || !view.activeLine) throw new Error('Expected text geometry');
  const entries = calls(canvas);
  let paintedText = '';
  const painted: { font: unknown; color: unknown; letterSpacing: unknown;
    baselineOffset: number; underline: boolean }[] = [];
  for (const [index, call] of entries.entries()) {
    if (call.kind !== 'fillText' || !call.style) continue;
    paintedText += String(call.args[0]);
    const size = Number(String(call.style.font).match(/([\d.]+)px/)?.[1]);
    const next = entries[index + 1];
    const style = { ...call.style, baselineOffset: view.activeLine.line.baseline - Number(call.args[2]),
      underline: next?.kind === 'fillRect' && next.args[1] === view.activeLine.line.baseline + size * 0.08 };
    for (let offset = 0; offset < String(call.args[0]).length; offset += 1) painted.push(style);
  }
  expect(paintedText.slice(start - view.activeLine.line.start, start - view.activeLine.line.start + text.length)).toBe(text);
  for (let offset = 0; offset < text.length; offset += 1) {
    const run = box.lines.flatMap((line) => line.runs)
      .find((run) => run.start <= start + offset && start + offset < run.end);
    expect(run).toBeDefined();
    if (!run) throw new Error('Expected authoritative inserted character');
    expect(painted[start - view.activeLine.line.start + offset]).toEqual({
      font: `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.fontSizePx}px ${JSON.stringify(run.fontFamily)}`,
      color: run.color, letterSpacing: `${run.letterSpacingPx ?? 0}px`,
      baselineOffset: run.baselineOffsetPx ?? 0, underline: run.underline,
    });
  }
}

async function checkEchoStyle(
  bytes: Uint8Array, position: number, kind: 'insertText' | 'replaceText' | 'compositionCommit' = 'insertText',
  focus = position, text = 'X', style?: pptx.TextStyle,
): Promise<void> {
  const value = await realSession('alpha', position, focus, 1, false, bytes);
  const { canvas, store, dispose } = echoView(value.client);
  try {
    if (kind === 'compositionCommit') store.beginComposition('echo-style-ime', 100);
    calls(canvas).length = 0;
    let pending: Promise<PresentationCommandAck | null>;
    if (kind === 'compositionCommit') {
      store.updateComposition(text);
      const preview = store.getSnapshot();
      const recorded = [...calls(canvas)];
      pending = commitComposition(store, text, 100);
      const ack = await pending;
      expect(ack?.outcomes.some((outcome) => outcome.committed)).toBe(true);
      expect((await value.savedParagraphs()).join('\n').slice(Math.min(position, focus), Math.min(position, focus) + text.length)).toBe(text);
      calls(canvas).splice(0, calls(canvas).length, ...recorded);
      expectEchoStyle(canvas, preview, ack!.frame!, Math.min(position, focus), text);
    } else {
      const at = { source: 'confirmed' as const, selection: { ...textSelection(store.getSnapshot().selection), anchor: position, focus } };
      pending = store.dispatch({ kind, payload: { at, text, style, inputTime: 100, undoGroup: 'echo-style' } });
      const echo = store.getSnapshot();
      const recorded = [...calls(canvas)];
      const ack = await pending;
      expect(ack?.outcomes.some((outcome) => outcome.committed)).toBe(true);
      expect((await value.savedParagraphs()).join('\n').slice(Math.min(position, focus), Math.min(position, focus) + text.length)).toBe(text);
      calls(canvas).splice(0, calls(canvas).length, ...recorded);
      expectEchoStyle(canvas, echo, ack!.frame!, Math.min(position, focus), text);
    }
  } finally { dispose(); await value.dispose(); }
}

for (const kind of ['insertText', 'replaceText', 'compositionCommit'] as const) {
  test(`echo_style_${kind}_styled_deck`, async () => {
    for (const size of [42, 24]) await checkEchoStyle(await echoStyleDeck(styledEchoParagraph(size)), 5, kind);
    await checkEchoStyle(await echoStyleDeck(styledEchoParagraph()), 5, kind, 0, 'X\nY');
  });
}

test('echo_style_payload_marks_and_subsequent_unstyled_typing', async () => {
  await checkEchoStyle(await echoStyleDeck(styledEchoParagraph()), 0, 'insertText', 0, 'M', { bold: true });
  const value = await realSession('alpha', 0, 0, 1, false, await echoStyleDeck(styledEchoParagraph()));
  const { canvas, store, dispose } = echoView(value.client);
  try {
    await store.dispatch({ kind: 'insertText', payload: { at: { source: 'confirmed', selection: textSelection(value.client.state.selection) },
      text: 'M', style: { bold: true }, inputTime: 100, undoGroup: 'marks' } });
    await store.dispatch({ kind: 'setSelection', payload: { selection: { ...textSelection(value.client.state.selection), anchor: 6, focus: 6 },
      inputTime: 101, undoGroup: 'selection' } });
    calls(canvas).length = 0;
    const pending = store.dispatch({ kind: 'insertText', payload: { at: { source: 'confirmed', selection: textSelection(value.client.state.selection) },
      text: 'X', inputTime: 102, undoGroup: 'typing' } });
    const echo = store.getSnapshot();
    const recorded = [...calls(canvas)];
    const ack = await pending;
    calls(canvas).splice(0, calls(canvas).length, ...recorded);
    expectEchoStyle(canvas, echo, ack.frame!, 6, 'X');
    const box = pptx.decodeCommandDisplayList(ack.frame!).primitives.find((primitive) => primitive.kind === 'textBox');
    expect(box?.kind === 'textBox' && box.lines.flatMap((line) => line.runs)
      .find((run) => run.start <= 6 && run.end > 6)?.bold).toBe(false);
  } finally { dispose(); await value.dispose(); }
});

function echoRun(text: string, attributes: string): string {
  return `<a:r><a:rPr sz="2400" u="none" ${attributes}><a:solidFill><a:srgbClr val="3478CC"/></a:solidFill><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${text}</a:t></a:r>`;
}

test('echo_style_differently_styled_run_boundary', async () => {
  await checkEchoStyle(await echoStyleDeck(`<a:p>${echoRun('a', 'b="1" i="0"')}${echoRun('b', 'b="0" i="1"')}</a:p>`), 1);
});

test('echo_style_repeated_characters_use_caret_offset', async () => {
  await checkEchoStyle(await echoStyleDeck(`<a:p>${echoRun('a', 'b="1" i="0"')}${echoRun('a', 'b="0" i="1"')}</a:p>`), 0,
    'insertText', 0, 'a');
});

test('echo_style_inherited_size_differs_from_editor_fallback', async () => {
  await checkEchoStyle(await echoStyleDeck(styledEchoParagraph().replace('sz="4200"', '')), 5);
});

test('echo_style_spacing_baseline_and_caps_are_not_inherited_by_insert', async () => {
  for (const attributes of ['spc="250"', 'baseline="30000"', 'cap="small"']) {
    await checkEchoStyle(await echoStyleDeck(`<a:p>${echoRun('alpha', `b="0" i="0" ${attributes}`)}</a:p>`), 5);
  }
});

test('echo_style_worker_created_run_keeps_certainty_for_chained_inserts', async () => {
  const value = await realSession('alpha', 5, 5, 1, false, await echoStyleDeck(styledEchoParagraph()));
  const { canvas, painter, store, dispose } = echoView(value.client);
  try {
    await store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'a',
      style: { bold: true, italic: true, underline: 'none', fontSizePt: 24,
        color: '#3478CC', fontFamily: 'Liberation Sans' }, inputTime: 100, undoGroup: 'setup',
    } });
    calls(canvas).length = 0;
    const held = value.hold();
    const first = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'a',
      inputTime: 101, undoGroup: 'typing',
    } });
    const second = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'afterSeq', seq: value.client.state.throughSeq + 1 }, text: 'a',
      inputTime: 102, undoGroup: 'typing',
    } });
    const echo = store.getSnapshot();
    expectNoEcho(canvas, echo);
    expect(echo.pending.every((record) => !record.echoed)).toBe(true);
    expect(echo.waitingInputCount).toBe(1);
    calls(canvas).length = 0;
    painter.paintEcho(echo);
    const recorded = [...calls(canvas)];
    await held;
    value.release();
    const [, ack] = await Promise.all([first, second]);
    calls(canvas).splice(0, calls(canvas).length, ...recorded);
    expectEchoStyle(canvas, echo, ack.frame!, 6, 'aa');
    expect(frameRuns(ack.frame!, textSelection(value.client.state.selection).storyId).map((run) => run.text).join('')).toBe('alphaaaa');
    const inserted = frameRuns(ack.frame!, textSelection(value.client.state.selection).storyId).filter((run) => run.start < 8 && run.end > 6);
    expect(inserted.length).toBeGreaterThan(0);
    expect(inserted.every((run) => run.bold && run.italic && !run.underline
      && run.fontSizePx === 32 && run.color === '#3478CC')).toBe(true);
  } finally { dispose(); await value.dispose(); }
});

test('echo_allowlist_arabic_coverage_fallback', async () => {
  const value = await realSession('abc', 3, 3, 1, false, undefined, true);
  const { canvas, store, dispose } = echoView(value.client);
  try {
    expect(value.client.state.editState.toolbar.echoRun).toBe('before');
    const runs = store.getSnapshot().activeLine!.line.runs;
    const neighbour = runs[runs.length - 1]!;
    calls(canvas).length = 0;
    const pending = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'ش', inputTime: 100, undoGroup: 'typing',
    } });
    expectNoEcho(canvas, store.getSnapshot());
    const pendingIntents = store.getSnapshot().pending;
    expect(pendingIntents[pendingIntents.length - 1]?.echoed).toBe(false);
    const ack = await pending;
    const inserted = frameRuns(ack.frame!, textSelection(value.client.state.selection).storyId).find((run) => run.text.includes('ش'));
    expect(inserted?.fontFamily).toBe('Noto Sans Arabic');
    expect(inserted?.fontId).not.toBe(neighbour.fontId);
    expect(await value.savedParagraphs()).toEqual(['abcش']);
  } finally { dispose(); await value.dispose(); }
});

test('echo_allowlist_inherited_small_caps', async () => {
  const bytes = await echoStyleDeck(`<a:p><a:pPr><a:defRPr cap="small"/></a:pPr>${echoRun('alpha', 'b="0" i="0"')}</a:p>`);
  const value = await realSession('alpha', 5, 5, 1, false, bytes);
  const { canvas, store, dispose } = echoView(value.client);
  try {
    await store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'A', inputTime: 100, undoGroup: 'setup',
    } });
    expect(value.client.state.editState.toolbar.echoRun).toBeNull();
    const runs = store.getSnapshot().activeLine!.line.runs;
    const neighbour = runs[runs.length - 1]!;
    expect(neighbour.text).toBe('A');
    calls(canvas).length = 0;
    const pending = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'a', inputTime: 101, undoGroup: 'typing',
    } });
    expectNoEcho(canvas, store.getSnapshot());
    const pendingIntents = store.getSnapshot().pending;
    expect(pendingIntents[pendingIntents.length - 1]?.echoed).toBe(false);
    const ack = await pending;
    const inserted = frameRuns(ack.frame!, textSelection(value.client.state.selection).storyId).find((run) => run.start <= 6 && run.end > 6)!;
    expect(inserted.text).toContain('A');
    expect(inserted.fontSizePx).toBeLessThan(neighbour.fontSizePx);
    expect(await value.savedParagraphs()).toEqual(['alphaAa']);
  } finally { dispose(); await value.dispose(); }
});

test('echo_allowlist_latin1_requires_confirmed_font_coverage', async () => {
  const value = await realSession('abc', 3);
  const { canvas, store, dispose } = echoView(value.client);
  try {
    expect(value.client.state.editState.toolbar.echoRun).toBe('before');
    calls(canvas).length = 0;
    const first = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'é', inputTime: 100, undoGroup: 'typing',
    } });
    expectNoEcho(canvas, store.getSnapshot());
    const pendingIntents = store.getSnapshot().pending;
    expect(pendingIntents[pendingIntents.length - 1]?.echoed).toBe(false);
    await first;
    expect(await value.savedParagraphs()).toEqual(['abcé']);
    expect(store.getSnapshot().activeLine!.text).toBe('abcé');
    calls(canvas).length = 0;
    const second = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'é', inputTime: 101, undoGroup: 'typing',
    } });
    const echo = store.getSnapshot();
    expectNoEcho(canvas, echo);
    expect(echo.pending[echo.pending.length - 1]?.echoed).toBe(false);
    const recorded = [...calls(canvas)];
    const ack = await second;
    calls(canvas).splice(0, calls(canvas).length, ...recorded);
    expectEchoStyle(canvas, echo, ack.frame!, 4, 'é');
    expect(await value.savedParagraphs()).toEqual(['abcéé']);
  } finally { dispose(); await value.dispose(); }
});

test('echo_allowlist_inherited_spacing_and_baseline', async () => {
  for (const [attribute, property] of [['spc="250"', 'letterSpacingPx'], ['baseline="30000"', 'baselineOffsetPx']] as const) {
    const bytes = await echoStyleDeck(`<a:p><a:pPr><a:defRPr ${attribute}/></a:pPr>${echoRun('alpha', 'b="0" i="0"')}</a:p>`);
    const value = await realSession('alpha', 5, 5, 1, false, bytes);
    const { canvas, store, dispose } = echoView(value.client);
    try {
      await store.dispatch({ kind: 'insertText', payload: {
        at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'a', inputTime: 100, undoGroup: 'setup',
      } });
      expect(value.client.state.editState.toolbar.echoRun).toBeNull();
      const runs = store.getSnapshot().activeLine!.line.runs;
      expect(runs[runs.length - 1]![property]).toBeGreaterThan(0);
      calls(canvas).length = 0;
      const pending = store.dispatch({ kind: 'insertText', payload: {
        at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'a', inputTime: 101, undoGroup: 'typing',
      } });
      expectNoEcho(canvas, store.getSnapshot());
      const pendingIntents = store.getSnapshot().pending;
      expect(pendingIntents[pendingIntents.length - 1]?.echoed).toBe(false);
      const ack = await pending;
      expect(frameRuns(ack.frame!, textSelection(value.client.state.selection).storyId).map((run) => run.text).join('')).toBe('alphaaa');
      expect(await value.savedParagraphs()).toEqual(['alphaaa']);
    } finally { dispose(); await value.dispose(); }
  }
});

test('echo_allowlist_composition_preview_uses_confirmed_coverage', async () => {
  const value = await realSession('abc', 3);
  const { canvas, store, dispose } = echoView(value.client);
  try {
    store.beginComposition('allowlist-ime', 100);
    await store.fence();
    for (const text of ['ش', 'é', 'a\nY', 'a\t', 'a漢']) {
      calls(canvas).length = 0;
      store.updateComposition(text);
      expectNoEcho(canvas, store.getSnapshot());
    }
    calls(canvas).length = 0;
    store.updateComposition('a');
    const echo = store.getSnapshot();
    expectNoEcho(canvas, echo);
    const recorded = [...calls(canvas)];
    const ack = await commitComposition(store, 'a', 101);
    calls(canvas).splice(0, calls(canvas).length, ...recorded);
    expectEchoStyle(canvas, echo, ack!.frame!, 3, 'a');
    expect(await value.savedParagraphs()).toEqual(['abca']);
  } finally { dispose(); await value.dispose(); }
});

test('echo_allowlist_composition_preview_rejects_inherited_transforms', async () => {
  for (const attribute of ['cap="small"', 'spc="250"', 'baseline="30000"']) {
    const bytes = await echoStyleDeck(`<a:p><a:pPr><a:defRPr ${attribute}/></a:pPr>${echoRun('alpha', 'b="0" i="0"')}</a:p>`);
    const value = await realSession('alpha', 5, 5, 1, false, bytes);
    const { canvas, store, dispose } = echoView(value.client);
    try {
      await store.dispatch({ kind: 'insertText', payload: {
        at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'A', inputTime: 100, undoGroup: 'setup',
      } });
      expect(value.client.state.editState.toolbar.echoRun).toBeNull();
      store.beginComposition('inherited-ime', 101);
      await store.fence();
      calls(canvas).length = 0;
      store.updateComposition('a');
      expectNoEcho(canvas, store.getSnapshot());
      const ack = await commitComposition(store, 'a', 102);
      expect(ack?.outcomes.some((outcome) => outcome.committed)).toBe(true);
      expect(await value.savedParagraphs()).toEqual(['alphaAa']);
    } finally { dispose(); await value.dispose(); }
  }
});

async function realSession(text = 'abc', position = 1, focus = position, slideCount = 1, fields = false,
  bytes?: Uint8Array, arabicFallback = false) {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
    handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
  });
  const commandIds = new Set<number>();
  const sent: PresentationCommandEnvelope[] = [];
  const held: unknown[] = [];
  let deliver: ((message: unknown) => void) | undefined;
  let holding = false;
  let reached: (() => void) | undefined;
  let heldCount = 1;
  const transport: SessionTransport = { ...pair.client,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'command') {
        commandIds.add(message.id);
        sent.push(message.args[0] as PresentationCommandEnvelope);
      }
      pair.client.post(message, transfer);
    },
    listen(listener) {
      deliver = listener;
      return pair.client.listen((message) => {
        if (holding && isHostMessage(message) && message.kind === 'reply' && commandIds.has(message.id)) {
          held.push(message);
          if (held.length >= heldCount) reached?.();
        } else listener(message);
      });
    },
  };
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const fallbackFonts = arabicFallback ? [{ family: 'Noto Sans Arabic', bytes: new Uint8Array(await readFile(
    new URL('../../../fonts/assets/NotoSansArabic-Regular.ttf', import.meta.url))).buffer }] : [];
  const client = await createPresentationCommandClient(bytes ?? await expandedDeck(text, slideCount, fields), {
    session: 'p4-session', incarnation: 'p4-worker', clientId: 7822,
    fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }], fallbackFonts,
  }, transport);
  const content = (await client.read()).result;
  if (!content.ok || !content.stories[0]) throw new Error('Expected text story');
  const story = content.stories[0];
  if (bytes === undefined && !fields) {
    await client.dispatch({ kind: 'replaceText', payload: {
      at: { source: 'confirmed', selection: { slideId: story.slideId, shapeId: story.shapeId,
        storyId: story.storyId, anchor: 0, focus: text.length, affinity: 'downstream' } },
      text, style: { bold: false, italic: false, underline: 'none', fontSizePt: 12,
        color: '#000000', fontFamily: 'Liberation Sans' }, inputTime: 0, undoGroup: 'setup',
    } });
  }
  await client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
      anchor: position, focus, affinity: 'downstream' }, inputTime: 0, undoGroup: 'selection',
  } });
  const outcomes = new Map<number, PresentationCommandAck['outcomes'][number]>();
  const unsubscribe = client.subscribe((ack) => {
    for (const outcome of ack.outcomes) outcomes.set(outcome.seq, outcome);
  });
  const owner: PeerlessSession = {
    ready: Promise.resolve(client), failure: undefined,
    subscribeFailure: (listener) => client.subscribeFailure(listener), dispose: mock(async () => {}),
  };
  const opener = spyOn(peerlessSessionOpener, 'open').mockReturnValue(owner);
  restorers.push(() => opener.mockRestore());
  return { client, sent, outcomes, start: sent.length,
    hold(count = 1) {
      holding = true;
      heldCount = count;
      return new Promise<void>((resolve) => { reached = resolve; });
    },
    release() {
      holding = false;
      for (const reply of held.splice(0)) deliver?.(reply);
    },
    releaseFirst() {
      const reply = held.shift();
      if (reply === undefined) throw new Error('Expected held reply');
      deliver?.(reply);
    },
    async savedParagraphs() {
      const zip = await JSZip.loadAsync(await client.save());
      const slide = zip.file('ppt/slides/slide1.xml');
      if (!slide) throw new Error('Expected saved slide');
      const source = await slide.async('string');
      return [...source.matchAll(/<a:p(?:\s[^>]*)?>([\s\S]*?)<\/a:p>/g)].map((paragraph) =>
        [...paragraph[1].matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)].map((run) =>
          run[1].replace(/&(?:amp|lt|gt|quot|apos);/g, (entity) =>
            ({ '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" } as Record<string, string>)[entity] ?? entity)).join(''));
    },
    async dispose() {
      unsubscribe();
      try { await client.dispose(); } finally { registration.dispose(); pair.client.close(); }
    },
  };
}

async function mountedRealSession(text = 'abc', position = 1, focus = position, slideCount = 1, fields = false) {
  const value = await realSession(text, position, focus, slideCount, fields);
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const errors = mock((_error: Error) => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} showToolbar={false}
    onReady={ready} onError={errors} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const input = view.container.querySelector('textarea');
  if (!input) throw new Error('Expected peer-less text input');
  const canvas = view.getByTestId('pptx-slide-canvas');
  if (!(canvas instanceof HTMLCanvasElement)) throw new Error('Expected slide canvas');
  const confirmed = paints[0];
  if (!confirmed) throw new Error('Expected installed frame');
  const { width, height } = confirmed.displayList;
  const bounds = new DOMRect(100, 50, width * 2, height * 2);
  const rect = spyOn(canvas, 'getBoundingClientRect').mockReturnValue(bounds);
  restorers.push(() => rect.mockRestore());
  return { ...value, input, canvas, confirmed, errors,
    flush: () => ready.mock.calls[0][0].flushPendingInput(),
    click(point: { x: number; y: number }) {
      const pointer = { pointerId: 1, isPrimary: true, button: 0,
        clientX: bounds.left + point.x * 2, clientY: bounds.top + point.y * 2 };
      fireEvent.pointerDown(canvas, pointer);
      fireEvent.pointerUp(canvas, pointer);
    },
    async finish(expected: string, kinds: readonly (string | readonly string[])[]) {
      await act(async () => { value.release(); await ready.mock.calls[0][0].flushPendingInput(); });
      const intents = value.sent.slice(value.start);
      expect(intents).toHaveLength(kinds.length);
      expect(value.outcomes.size).toBe(kinds.length);
      for (const [index, intent] of intents.entries()) {
        const kind = kinds[index];
        expect(typeof kind === 'string' ? [kind] : kind).toContain(intent.kind);
        const outcome = value.outcomes.get(intent.seq);
        expect(outcome?.status === 'applied' || outcome?.status === 'noop').toBe(true);
        if (['insertText', 'replaceText', 'deleteBackward'].includes(intent.kind)) expect(outcome?.committed).toBe(true);
        if (index === 0) continue;
        const predecessor = intents[index - 1].seq;
        expect(intent.afterSeq).toBe(predecessor);
        expect('at' in intent.payload ? intent.payload.at : undefined).toEqual({ source: 'afterSeq', seq: predecessor });
      }
      expect(await value.savedParagraphs()).toEqual([expected]);
      expect(errors).not.toHaveBeenCalled();
    },
    async dispose() { value.release(); view.unmount(); await value.dispose(); },
  };
}

function beforeInput(input: HTMLTextAreaElement, inputType: string, data: string | null = null, inputTime?: number) {
  const event = new Event('beforeinput', { bubbles: true, cancelable: true });
  Object.defineProperties(event, { inputType: { value: inputType }, data: { value: data },
    ...(inputTime === undefined ? {} : { timeStamp: { value: inputTime } }) });
  fireEvent(input, event);
  expect(event.defaultPrevented).toBe(true);
}

function compose(input: HTMLTextAreaElement, type: string, data: string): void {
  const event = new Event(type, { bubbles: true });
  Object.defineProperty(event, 'data', { value: data });
  fireEvent(input, event);
}

function framePoint(frame: PeerlessConfirmedFrame, position: number, pastEnd = false) {
  const box = frame.displayList.primitives.find((primitive) => primitive.kind === 'textBox');
  if (!box || box.kind !== 'textBox') throw new Error('Expected text box');
  const line = box.lines.find((candidate) => candidate.caretStops.some((stop) => stop.position === position));
  const stop = line?.caretStops.find((candidate) => candidate.position === position);
  if (!line || !stop) throw new Error('Expected caret stop');
  const x = stop.x + (pastEnd ? Math.min(20, (box.x + box.w - stop.x) / 2) : 0);
  if (pastEnd) expect(x).toBeGreaterThan(stop.x);
  return { x, y: line.y + line.height / 2 };
}

test('p4_click_past_end_then_left_resolves_after_held_insert', async () => {
  const value = await mountedRealSession();
  try {
    const held = value.hold();
    await act(async () => {
      beforeInput(value.input, 'insertText', 'X');
      value.click(framePoint(value.confirmed, 3, true));
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      beforeInput(value.input, 'insertText', 'Y');
    });
    await held;
    await value.finish('aXbYc', ['insertText', 'selectAtPoint', 'moveSelection', 'insertText']);
    expect(value.sent[value.start + 2]).toMatchObject({ payload: { motion: 'left', extend: false } });
  } finally { await value.dispose(); }
});

test('p4_end_type_left_type_preserves_queued_order', async () => {
  const value = await mountedRealSession('abc', 3);
  try {
    const held = value.hold();
    await act(async () => {
      beforeInput(value.input, 'insertText', 'X');
      fireEvent.keyDown(value.input, { key: 'End' });
      beforeInput(value.input, 'insertText', 'Y');
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      beforeInput(value.input, 'insertText', 'Z');
    });
    await held;
    await value.finish('abcXZY', ['insertText', 'moveSelection', 'insertText', 'moveSelection', 'insertText']);
    expect(value.sent[value.start + 1]).toMatchObject({ payload: { motion: 'lineEnd', extend: false } });
    expect(value.sent[value.start + 3]).toMatchObject({ payload: { motion: 'left', extend: false } });
  } finally { await value.dispose(); }
});

test('p4_paragraph_join_end_type_survives_first_deletion_reply', async () => {
  const value = await mountedRealSession('a\nbc', 2);
  try {
    expect(await value.savedParagraphs()).toEqual(['a', 'bc']);
    const held = value.hold();
    await act(async () => {
      beforeInput(value.input, 'deleteContentBackward');
      fireEvent.keyDown(value.input, { key: 'End' });
      beforeInput(value.input, 'insertText', 'Y');
    });
    await held;
    await act(async () => { value.release(); await value.flush(); });
    const intents = value.sent.slice(value.start);
    expect(intents.map((intent) => intent.kind)).toEqual(['deleteBackward', 'moveSelection', 'insertText']);
    expect(value.outcomes.get(intents[0].seq)).toMatchObject({ status: 'applied', committed: true });
    expect(value.outcomes.get(intents[1].seq)).toMatchObject({ status: 'applied', committed: false });
    expect(value.outcomes.get(intents[2].seq)).toMatchObject({ status: 'applied', committed: true });
    for (const [index, intent] of intents.slice(1).entries()) {
      expect(intent.afterSeq).toBe(intents[index].seq);
      expect('at' in intent.payload ? intent.payload.at : undefined).toEqual({ source: 'afterSeq', seq: intents[index].seq });
    }
    expect(await value.savedParagraphs()).toEqual(['abcY']);
    expect(value.errors).not.toHaveBeenCalled();
    expect(value.sent[value.start + 1]).toMatchObject({ payload: { motion: 'lineEnd', extend: false } });
  } finally { await value.dispose(); }
});

test('p4_oversized_paste_shift_end_replaces_only_suffix', async () => {
  const value = await mountedRealSession('abc', 1, 2);
  try {
    const held = value.hold();
    const paste = 'X'.repeat(8193);
    await act(async () => {
      beforeInput(value.input, 'insertFromPaste', paste);
      fireEvent.keyDown(value.input, { key: 'End', shiftKey: true });
      beforeInput(value.input, 'insertText', 'Y');
    });
    await held;
    await value.finish(`a${paste}Y`, ['replaceText', 'moveSelection', ['insertText', 'replaceText']]);
    expect(value.sent[value.start + 1]).toMatchObject({ payload: { motion: 'lineEnd', extend: true } });
    const final = paints[paints.length - 1];
    if (!final) throw new Error('Expected covering frame');
    const box = final.displayList.primitives.find((primitive) => primitive.kind === 'textBox');
    if (!box || box.kind !== 'textBox') throw new Error('Expected text box');
    expect(box.lines).toHaveLength(1);
  } finally { await value.dispose(); }
});

test('p4_held_click_sends_frame_point_and_chains_typing', async () => {
  const value = await mountedRealSession('abc', 3);
  try {
    const point = framePoint(value.confirmed, 1);
    const held = value.hold();
    await act(async () => {
      beforeInput(value.input, 'insertText', 'X');
      value.click(point);
      beforeInput(value.input, 'insertText', 'Y');
    });
    await held;
    await value.finish('aYbcX', ['insertText', 'selectAtPoint', 'insertText']);
    const click = value.sent[value.start + 1];
    const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = value.confirmed.frame;
    expect(click).toMatchObject({ kind: 'selectAtPoint', payload: {
      frame: { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation },
      extend: false,
    } });
    if (click.kind !== 'selectAtPoint') throw new Error('Expected frame-point click');
    expect(click.payload.x).toBeCloseTo(point.x, 6);
    expect(click.payload.y).toBeCloseTo(point.y, 6);
    expect(Object.keys(click.payload).sort()).toEqual(['at', 'drag', 'extend', 'frame', 'granularity', 'inputTime', 'undoGroup', 'x', 'y']);
    expect(click.payload).toMatchObject({ drag: false, granularity: 'character' });
    expect(click.payload).not.toHaveProperty('selection');
    expect(click.payload).not.toHaveProperty('anchor');
    expect(click.payload).not.toHaveProperty('focus');
    expect(value.sent[value.start + 2]).toMatchObject({ afterSeq: click.seq,
      payload: { at: { source: 'afterSeq', seq: click.seq }, text: 'Y' } });
  } finally { await value.dispose(); }
});

test('p4_echo_stops_before_worker_resolved_navigation', async () => {
  const value = await mountedRealSession('abc', 3);
  try {
    const held = value.hold();
    await act(async () => { beforeInput(value.input, 'insertText', 'a'); });
    expect(kinds(calls(value.canvas).filter((entry) => entry.kind === 'fillText'))).toEqual([]);
    calls(value.canvas).length = 0;
    await act(async () => {
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      beforeInput(value.input, 'insertText', 'b');
    });
    await held;
    expect(paints[paints.length - 1]?.frame.coversThroughSeq).toBe(value.confirmed.frame.coversThroughSeq);
    expect(calls(value.canvas).some((entry) => entry.kind === 'fillText'
      && typeof entry.args[0] === 'string' && entry.args[0].length > 4)).toBe(false);
    await value.finish('abcba', ['insertText', 'moveSelection', 'insertText']);
    expect(value.sent[value.start + 1]).toMatchObject({ payload: { motion: 'left', extend: false } });
  } finally { await value.dispose(); }
});

test('p5_blank_click_then_text_click_keeps_held_typing', async () => {
  const value = await mountedRealSession('abc', 1);
  try {
    const held = value.hold();
    await act(async () => {
      beforeInput(value.input, 'insertText', 'X');
      value.click({ x: value.confirmed.displayList.width - 1, y: value.confirmed.displayList.height - 1 });
      value.click(framePoint(value.confirmed, 3, true));
      beforeInput(value.input, 'insertText', 'Y');
    });
    await held;
    await value.finish('aXbcY', ['insertText', 'setSelection', 'selectAtPoint', 'insertText']);
    expect(value.sent[value.start + 1]).toMatchObject({ payload: { selection: null } });
  } finally { await value.dispose(); }
});

test.each(['X', 'XY'])('p5_blank_click_keeps_held_typing_without_a_followup (%s)', async (text) => {
  const value = await mountedRealSession('abc', 1);
  try {
    const held = value.hold();
    await act(async () => {
      for (const character of text) beforeInput(value.input, 'insertText', character);
      value.click({ x: value.confirmed.displayList.width - 1, y: value.confirmed.displayList.height - 1 });
    });
    await held;
    expect(value.sent.slice(value.start).map((intent) => intent.kind)).toEqual(['insertText']);
    await value.finish(`a${text}bc`, [...Array.from(text, () => 'insertText'), 'setSelection']);
    expect(value.sent[value.start + text.length]).toMatchObject({ payload: { selection: null } });
    expect(value.client.state.selection).toBeNull();
    expect(paintedViews[paintedViews.length - 1]?.selection).toBeNull();
  } finally { await value.dispose(); }
});

test('p5_capped_typing_navigation_survives_expired_basis', async () => {
  const value = await mountedRealSession('abc', 1);
  try {
    const held = value.hold();
    await act(async () => {
      for (let index = 0; index < 256; index += 1) beforeInput(value.input, 'insertText', 'a', 100);
      beforeInput(value.input, 'insertText', 'y', 101);
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      beforeInput(value.input, 'insertText', 'Z');
    });
    await held;
    expect(value.sent.slice(value.start)).toHaveLength(1);
    expect(kinds(calls(value.canvas).filter((entry) => entry.kind === 'fillText'))).toEqual([]);
    await value.finish(`a${'a'.repeat(256)}Zybc`, ['insertText', 'insertText', 'insertText', 'moveSelection', 'insertText']);
  } finally { await value.dispose(); }
}, 30_000);

test('p6_capped_typing_navigation_preserves_queued_click_frame', async () => {
  const value = await mountedRealSession('abc', 3);
  try {
    const held = value.hold();
    await act(async () => {
      for (let index = 0; index < 256; index += 1) beforeInput(value.input, 'insertText', 'a', 100);
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      value.click(framePoint(value.confirmed, 1));
      beforeInput(value.input, 'insertText', 'Z');
    });
    await held;
    expect(value.sent.slice(value.start)).toHaveLength(1);
    const typed = value.hold();
    await act(async () => { value.releaseFirst(); });
    await typed;
    expect(value.sent.slice(value.start).map((intent): string => intent.kind)).toEqual([
      'insertText', 'insertText',
    ]);
    await value.finish(`aZbc${'a'.repeat(256)}`, [
      'insertText', 'insertText', 'moveSelection', 'selectAtPoint', 'insertText',
    ]);
  } finally { await value.dispose(); }
}, 30_000);

test('p7_capped_typing_two_navigations_preserve_queued_click_frame', async () => {
  const value = await mountedRealSession('abc', 3);
  try {
    const held = value.hold();
    await act(async () => {
      for (let index = 0; index < 256; index += 1) beforeInput(value.input, 'insertText', 'a', 100);
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      value.click(framePoint(value.confirmed, 1));
      beforeInput(value.input, 'insertText', 'Z');
    });
    await held;
    expect(value.sent.slice(value.start)).toHaveLength(1);
    const typed = value.hold();
    await act(async () => { value.releaseFirst(); });
    await typed;
    expect(value.sent.slice(value.start).map((intent): string => intent.kind)).toEqual([
      'insertText', 'insertText',
    ]);
    const navigated = value.hold(4);
    await act(async () => { value.releaseFirst(); });
    await navigated;
    expect(value.sent.slice(value.start).map((intent): string => intent.kind)).toEqual([
      'insertText', 'insertText', 'moveSelection', 'moveSelection', 'selectAtPoint', 'insertText',
    ]);
    await value.finish(`aZbc${'a'.repeat(256)}`, [
      'insertText', 'insertText', 'moveSelection', 'moveSelection', 'selectAtPoint', 'insertText',
    ]);
  } finally { await value.dispose(); }
}, 30_000);

test('p5_thumbnails_preserve_displayed_frame_for_held_click', async () => {
  const value = await mountedRealSession('abc', 3, 3, 258);
  try {
    for (const slide of value.client.state.slides.slice(1)) await value.client.thumbnail(slide.id);
    const held = value.hold();
    await act(async () => {
      beforeInput(value.input, 'insertText', 'X');
      value.click(framePoint(value.confirmed, 1));
      beforeInput(value.input, 'insertText', 'Y');
    });
    await held;
    await value.finish('aYbcX', ['insertText', 'selectAtPoint', 'insertText']);
  } finally { await value.dispose(); }
}, 30_000);

test('p5_refused_field_commit_allows_later_plain_composition', async () => {
  const value = await mountedRealSession('A12B', 2, 2, 1, true);
  try {
    await act(async () => { compose(value.input, 'compositionstart', ''); });
    await waitFor(() => expect(value.outcomes.size).toBe(1));
    const held = value.hold();
    await act(async () => {
      compose(value.input, 'compositionupdate', '漢');
      compose(value.input, 'compositionend', '漢');
    });
    await held;
    await act(async () => { value.release(); await value.flush(); });
    const commit = value.sent.find((intent) => intent.kind === 'compositionCommit');
    if (!commit) throw new Error('Expected composition commit');
    expect(value.outcomes.get(commit.seq)).toMatchObject({ status: 'refused', reason: 'unsupported', committed: false });
    const plain = paints[paints.length - 1].displayList.primitives.find((primitive) =>
      primitive.kind === 'textBox' && primitive.lines.some((line) => line.runs.map((run) => run.text).join('') === 'abc'));
    if (plain?.kind !== 'textBox') throw new Error('Expected plain text box');
    const line = plain.lines[0];
    const stop = line.caretStops.find((candidate) => candidate.position === 3);
    if (!stop) throw new Error('Expected plain caret');
    await act(async () => { value.click({ x: stop.x, y: line.y + line.height / 2 }); });
    await act(async () => { await value.flush(); });
    await act(async () => {
      compose(value.input, 'compositionstart', '');
      compose(value.input, 'compositionupdate', '語');
      compose(value.input, 'compositionend', '語');
    });
    await act(async () => { await value.flush(); });
    expect(await value.savedParagraphs()).toEqual(['A12B', 'abc語']);
    expect(value.outcomes.get(value.sent[value.sent.length - 2].seq)).toMatchObject({ status: 'applied', committed: true });
    expect(value.errors.mock.calls.map(([error]) => error)).toMatchObject([
      { cause: { operation: 'compositionCommit', outcome: { reason: 'unsupported' } } },
      { cause: { operation: 'compositionEnd', outcome: { reason: 'dependency-refused' } } },
    ]);
  } finally { await value.dispose(); }
});

test('p8_capped_typing_two_displayed_clicks_preserve_later_input', async () => {
  const value = await mountedRealSession('abc', 3);
  try {
    const held = value.hold();
    await act(async () => {
      for (let index = 0; index < 256; index += 1) beforeInput(value.input, 'insertText', 'a', 100);
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      fireEvent.keyDown(value.input, { key: 'ArrowLeft' });
      value.click(framePoint(value.confirmed, 1));
    });
    await held;
    const typed = value.hold();
    await act(async () => { value.releaseFirst(); });
    await typed;
    const firstNavigated = value.hold(3);
    await act(async () => { value.releaseFirst(); });
    await firstNavigated;
    const displayed = paints[paints.length - 1];
    if (!displayed) throw new Error('Expected newly displayed frame');
    expect(displayed.frame.coversThroughSeq).toBe(value.sent[value.start + 1].seq);
    const secondNavigated = value.hold(5);
    await act(async () => {
      value.click(framePoint(displayed, 2));
      beforeInput(value.input, 'insertText', 'Z');
    });
    await secondNavigated;
    await value.finish(`abZc${'a'.repeat(256)}`, [
      'insertText', 'insertText', 'moveSelection', 'moveSelection',
      'selectAtPoint', 'selectAtPoint', 'insertText',
    ]);
  } finally { await value.dispose(); }
}, 30_000);

async function recSurfaceSession(echoReady = false) {
  const pair = createInProcessPair();
  let disposeHost = () => {};
  const hostTransport: SessionTransport = { ...pair.host, listen(listener) {
    disposeHost = () => listener({ protocol: 1, kind: 'dispose' });
    return pair.host.listen(listener);
  } };
  const sent: { id: number; method: string; envelope: unknown }[] = [];
  const methods: string[] = [];
  const held: Extract<HostMessage, { kind: 'reply' }>[] = [];
  const waiters = new Map<number, ReturnType<typeof deferred<void>>>();
  let deliver: ((message: unknown) => void) | undefined;
  let holding = false;
  let imageReplies = false;
  createPresentationSessionHost(hostTransport, { initWasm: async () => {
    await initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  const transport: SessionTransport = { ...pair.client,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call') {
        methods.push(message.method);
        if (message.method === 'command' || message.method === 'edit') {
          sent.push({ id: message.id, method: message.method, envelope: message.args[0] });
        }
      }
      pair.client.post(message, transfer);
    },
    listen(listener) {
      deliver = listener;
      return pair.client.listen((raw) => {
        let message = raw;
        if (isHostMessage(message) && message.kind === 'reply') {
          const id = message.id;
          if (!sent.some((call) => call.id === id)) { listener(message); return; }
          if (imageReplies && message.ok) {
            const ack = message.value;
            if (isCommandAck(ack) && ack.frame) {
              message = { ...message, value: { ...ack, frame: { ...ack.frame,
                payload: pictureFrame(ack.frame.payload as PresentationCommandFrame, 'tile') } } };
            } else if (ack !== null && typeof ack === 'object' && 'frame' in ack && ack.frame) {
              message = { ...message, value: { ...ack, frame: pictureFrame(ack.frame as PresentationCommandFrame, 'tile') } };
            }
          }
          if (holding && isHostMessage(message) && message.kind === 'reply') {
            held.push(message);
            for (const [count, waiter] of waiters) if (held.length >= count) waiter.resolve();
            return;
          }
        }
        listener(message);
      });
    },
  };
  const events = new EventTarget();
  const off = transport.listen((message) => events.dispatchEvent(new MessageEvent('message', { data: message })));
  const worker = {
    postMessage(message: unknown, transfer?: Transferable[]) { transport.post(message, transfer); },
    addEventListener: events.addEventListener.bind(events),
    removeEventListener: events.removeEventListener.bind(events),
    terminate() { disposeHost(); off(); transport.close(); },
  } as unknown as Worker;
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const factory = mock(() => worker);
  const owner = pptx.createPeerlessSession(await deck('abc'), {
    worker: factory, clientId: 7822,
    fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }],
  });
  const client = await owner.ready;
  const content = (await client.read()).result;
  expect(content.ok).toBe(true);
  if (!content.ok) throw new Error(content.failure.message);
  const story = content.stories[0]!;
  if (echoReady) {
    await client.dispatch({ kind: 'replaceText', payload: {
      at: { source: 'confirmed', selection: { slideId: story.slideId, shapeId: story.shapeId,
        storyId: story.storyId, anchor: 0, focus: 3, affinity: 'downstream' } }, text: 'abc',
      style: { bold: false, italic: false, underline: 'none', fontSizePt: 12,
        color: '#000000', fontFamily: 'Liberation Sans' }, inputTime: 0, undoGroup: 'setup',
    } });
  }
  await client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
      anchor: 1, focus: 1, affinity: 'downstream' }, inputTime: 0, undoGroup: 'selection',
  } });
  if (echoReady) expect(client.state.editState.toolbar.echoRun).toBe('after');
  sent.length = 0;
  methods.length = 0;
  const opener = spyOn(peerlessSessionOpener, 'open').mockReturnValue(owner);
  restorers.push(() => opener.mockRestore());
  return { client, owner, sent, methods, factory, baseSeq: client.state.throughSeq,
    hold() { holding = true; },
    images() { imageReplies = true; client.state.frame = pictureFrame(client.state.frame!, 'tile'); },
    waitForReplies(count: number) {
      if (held.length >= count) return Promise.resolve();
      const waiter = deferred<void>();
      waiters.set(count, waiter);
      return waiter.promise;
    },
    releaseOne() { const message = held.shift(); if (message) deliver?.(message); },
    release() { holding = false; for (const message of held.splice(0)) deliver?.(message); },
    unknown(seq: number) {
      const reply = held.shift();
      expect(reply).toBeDefined();
      if (reply) deliver?.({ protocol: 1, kind: 'reply', id: reply.id, ok: false,
        error: { name: 'CommandOutcomeUnknown', message: 'Edit outcome is unknown', refusal: { seq } } });
    },
    async dispose() { holding = false; await owner.dispose(); },
  };
}

test('rec_shared_echo_off_install_creates_no_pixel_snapshot', async () => {
  const value = await recSurfaceSession();
  const canvas = document.createElement('canvas');
  const framePainter = createCommandFramePainter(canvas, { scale: 1 / 960 });
  const store = createPeerlessViewStore(value.client, {
    ...framePainter, paintEcho() {}, restoreLine() {},
  });
  try {
    await framePainter.whenPainted();
    const before = store.getSnapshot();
    const count = canvasCalls.size;
    const candidate = { ...before.confirmed!.frame, slideRevision: before.confirmed!.frame.slideRevision + 1 };
    expect(peerlessViewStore.installCoveringFrame(store, candidate)).toBe(true);
    expect(canvasCalls.size - count).toBe(1);
    expect([...canvasCalls.values()].reduce((total, entries) => total + entries.filter((entry) =>
      entry.kind === 'drawImage' && entry.args[0] === canvas).length, 0)).toBe(0);
  } finally { store.dispose(); framePainter.dispose(); await value.dispose(); }
});

test('rec_shared_echo_throw_restores_painters_and_cancels_rejected_repaints', async () => {
  const value = await recSurfaceSession(true);
  const canvas = document.createElement('canvas');
  const image = heldImage();
  const framePainter = createCommandFramePainter(canvas, { scale: 1 / 960 });
  const echoPainter = createLocalEchoPainter(canvas, { scale: 1 / 960 });
  let reject = false;
  let drewCandidate = false;
  let candidateRaster: ReturnType<typeof commandFrameRaster>;
  const store = createPeerlessViewStore(value.client, {
    checkpoint() {
      const restoreFrame = framePainter.checkpoint?.();
      const restoreEcho = echoPainter.checkpoint?.();
      return Object.assign(() => { restoreFrame?.(); restoreEcho?.(); }, {
        commit() { restoreFrame?.commit?.(); restoreEcho?.commit?.(); },
      });
    },
    paintFrame(frame, view) {
      framePainter.paintFrame(frame, view);
      if (reject) {
        candidateRaster = commandFrameRaster(frame, 1 / 960);
        drewCandidate = calls(canvas).some((entry) => entry.kind === 'drawImage' && entry.args[0] === candidateRaster?.canvas);
      }
    },
    paintEcho(view) {
      echoPainter.paintEcho(view);
      if (reject) throw new Error('Echo declined');
    },
    restoreLine: echoPainter.restoreLine,
  }, { echoEnabled: true });
  try {
    value.hold();
    const operation = store.dispatch({ kind: 'insertText', payload: {
      at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'a', inputTime: 100, undoGroup: 'typing',
    } });
    void operation.catch(() => {});
    await value.waitForReplies(1);
    await framePainter.whenPainted();
    const before = store.getSnapshot();
    expect(before.echo?.text).toBe('aabc');
    const width = canvas.width;
    const height = canvas.height;
    const completion = framePainter.whenPainted();
    const raster = commandFrameRaster(before.confirmed!, 1 / 960);
    if (!raster) throw new Error('Expected confirmed raster');
    const plate = spyOn(raster, 'plate');
    restorers.push(() => plate.mockRestore());
    const picture = pictureFrame(before.confirmed!.frame, 'rejected');
    const list = JSON.parse(new TextDecoder().decode(picture.displayList)) as SlideDisplayList;
    list.width *= 2;
    const candidate = { ...picture, coversThroughSeq: value.baseSeq + 1,
      slideRevision: picture.slideRevision + 1, epoch: picture.epoch + 1,
      displayList: new TextEncoder().encode(JSON.stringify(list)).buffer };
    reject = true;
    expect(peerlessViewStore.installCoveringFrame(store, candidate)).toBe(false);
    expect(drewCandidate).toBe(false);
    expect(candidateRaster && [candidateRaster.canvas.width, candidateRaster.canvas.height]).toEqual([0, 0]);
    expect(canvas.width).toBe(width);
    expect(canvas.height).toBe(height);
    expect(framePainter.whenPainted() === completion).toBe(true);
    expect(store.getSnapshot().confirmed === before.confirmed).toBe(true);
    expect(store.getSnapshot().presentedThroughSeq).toBe(before.presentedThroughSeq);
    expect(store.getSnapshot().caret).toEqual(before.caret);
    expect(store.getSnapshot().echo).toEqual(before.echo);
    const draws = calls(canvas).filter((entry) => entry.kind === 'drawImage');
    expect(draws[draws.length - 1]?.args[0] === candidateRaster?.canvas).toBe(false);
    echoPainter.restoreLine(before);
    expect(plate.mock.calls.length).toBe(0);
    const count = calls(canvas).length;
    const source = new Image();
    source.width = source.height = 1;
    image.image.resolve(source);
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
    expect(calls(canvas).length).toBe(count);
    reject = false;
    value.release();
    await operation;
    await framePainter.whenPainted();
    expect(store.getSnapshot().activeLine?.text).toBe('aabc');
  } finally { store.dispose(); echoPainter.dispose(); framePainter.dispose(); await value.dispose(); }
});

test('rec_shared_unknown_outcome_keeps_typed_seq_and_last_painted_frame', async () => {
  const value = await recSurfaceSession();
  try {
    value.hold();
    const errors = mock((_error: Error) => {});
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} onError={errors} />);
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const input = view.container.querySelector('textarea');
    expect(input === null).toBe(false);
    if (!input) throw new Error('Expected text input');
    const painted = paints[paints.length - 1];
    await act(async () => { beforeInput(input, 'insertText', 'X'); });
    await value.waitForReplies(1);
    await act(async () => { value.unknown(value.baseSeq + 1); });
    await waitFor(() => expect(errors).toHaveBeenCalledTimes(1));
    const error = errors.mock.calls[0][0];
    expect(error).toBeInstanceOf(pptx.PptxWorkerEditorFailedError);
    expect(error).toMatchObject({ cause: { name: 'CommandOutcomeUnknown', seq: value.baseSeq + 1,
      message: 'Edit outcome is unknown', refusal: { seq: value.baseSeq + 1 } } });
    if (error instanceof pptx.PptxWorkerEditorFailedError) expect(error.cause).toBeInstanceOf(pptx.CommandOutcomeUnknown);
    expect(view.getByRole('application').getAttribute('aria-readonly')).toBe('true');
    expect(paints[paints.length - 1]).toBe(painted);
    const api = ready.mock.calls[0][0];
    await expect(api.saveAsync()).rejects.toBe(error);
    await expect(api.flushPendingInput()).rejects.toBe(error);
    await expect(api.version()).rejects.toBe(error);
    await expect(api.readContent()).rejects.toBe(error);
    expect(value.factory).toHaveBeenCalledTimes(1);
    expect(value.methods).not.toContain('checkpoint');
    expect(isCommandEnvelope(value.sent[0]!.envelope)).toBe(true);
    expect(value.sent[0]!.method).toBe('edit');
  } finally { await value.dispose(); }
});

test('rec_shared_save_waits_for_admission_queue_and_ticket_acks', async () => {
  echoEnabled = true;
  const value = await recSurfaceSession(true);
  try {
    value.hold();
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const input = view.container.querySelector('textarea');
    expect(input === null).toBe(false);
    if (!input) throw new Error('Expected text input');
    await act(async () => {
      for (let index = 0; index < 256; index += 1) beforeInput(input, 'insertText', 'a');
      beforeInput(input, 'insertText', 'LAST', 100);
      beforeInput(input, 'insertText', 'NEXT', 100);
    });
    expect(value.sent.length).toBe(256);
    const api = ready.mock.calls[0][0];
    const operations = [api.saveAsync(), api.flushPendingInput(), api.version(), api.readContent()] as const;
    for (const operation of operations) void operation.catch(() => {});
    await act(async () => { await Promise.resolve(); });
    expect(value.methods).not.toContain('saveCommandEditor');
    expect(value.methods).not.toContain('flushCommandEditor');
    expect(value.methods).not.toContain('readEditor');
    await value.waitForReplies(256);
    await act(async () => { value.releaseOne(); });
    await value.waitForReplies(256);
    expect(value.sent.length).toBe(257);
    expect(value.sent[256]!.envelope).toMatchObject({ seq: value.baseSeq + 257,
      afterSeq: value.baseSeq + 256, payload: { text: 'LASTNEXT' } });
    expect(value.methods).not.toContain('saveCommandEditor');
    expect(value.methods).not.toContain('flushCommandEditor');
    expect(value.methods).not.toContain('readEditor');
    await act(async () => { value.release(); await Promise.all(operations); });
    const [saved, , version, content] = await Promise.all(operations);
    expect(typeof version).toBe('string');
    expect(content.ok).toBe(true);
    if (content.ok) expect(content.stories[0]!.text).toBe(`a${'a'.repeat(256)}LASTNEXTbc`);
    const zip = await JSZip.loadAsync(saved);
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
    const savedText = [...slide.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)].map((match) => match[1]).join('');
    expect(savedText).toBe(`a${'a'.repeat(256)}LASTNEXTbc`);
    expect(value.methods.filter((method) => method === 'saveCommandEditor')).toHaveLength(1);
    expect(value.methods.filter((method) => method === 'flushCommandEditor')).toHaveLength(2);
    expect(value.methods.filter((method) => method === 'readEditor')).toHaveLength(1);
    expect(isCommandEnvelope(value.sent[256]!.envelope)).toBe(true);
    expect(value.sent.every((call) => call.method === 'edit')).toBe(true);
  } finally { echoEnabled = false; await value.dispose(); }
}, 30_000);

test('rec_shared_ready_waits_for_image_completion_without_delaying_install', async () => {
  const value = await recSurfaceSession();
  try {
    value.images();
    value.hold();
    const image = heldImage();
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const changed = mock((_snapshot: pptx.DeckSnapshot) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} onChange={changed} />);
    await waitFor(() => expect(paints.length).toBeGreaterThan(0));
    expect(ready).not.toHaveBeenCalled();
    const input = view.container.querySelector('textarea');
    expect(input).toBeNull();
    let operation!: Promise<PresentationCommandAck>;
    await act(async () => {
      operation = value.client.dispatch({ kind: 'insertText', payload: {
        at: { source: 'confirmed', selection: textSelection(value.client.state.selection) }, text: 'X',
        inputTime: 100, undoGroup: 'typing',
      } });
    });
    await value.waitForReplies(1);
    await act(async () => { value.releaseOne(); await operation; });
    const installed = paints[paints.length - 1];
    expect(installed.frame.coversThroughSeq).toBe(value.baseSeq + 1);
    expect(paintedViews[paintedViews.length - 1].confirmed).toBe(installed);
    expect(paintedViews[paintedViews.length - 1]).toMatchObject({ pending: [], echo: null,
      presentedThroughSeq: value.baseSeq + 1 });
    expect(ready).not.toHaveBeenCalled();
    expect(changed).not.toHaveBeenCalled();
    expect(value.methods).not.toContain('readEditor');
    let source!: CanvasImageSource;
    await act(async () => { source = image.release(); });
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    await ready.mock.calls[0][0].whenHydrated();
    expect(paints[paints.length - 1]).toBe(installed);
    expect([...canvasCalls.values()].some((entries) => entries.some((entry) =>
      entry.kind === 'drawImage' && entry.args[0] === source))).toBe(true);
    await waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
    expect(isCommandEnvelope(value.sent[0]!.envelope)).toBe(true);
    expect(value.sent[0]!.method).toBe('edit');
  } finally { await value.dispose(); }
});

describe('peerless caret blink', () => {
  let timers: ReturnType<typeof intervalClock>;
  let echoViews: PeerlessViewSnapshot[];
  let echoRestores: number;

  function intervalClock() {
    const originalSetInterval = window.setInterval;
    const originalClearInterval = window.clearInterval;
    let now = 0;
    let nextId = 0;
    const active = new Map<number, { tick(): void; delay: number; due: number }>();
    const started: number[] = [];
    const cleared: number[] = [];
    window.setInterval = ((tick: (...args: unknown[]) => void, delay?: number, ...args: unknown[]) => {
      if (delay !== 530) return originalSetInterval.call(window, tick, delay, ...args);
      const id = --nextId;
      started.push(id);
      active.set(id, { tick: () => tick(...args), delay, due: now + delay });
      return id;
    }) as typeof window.setInterval;
    window.clearInterval = ((id?: number) => {
      if (id === undefined) return;
      if (!started.includes(id)) { originalClearInterval.call(window, id); return; }
      cleared.push(id);
      active.delete(id);
    }) as typeof window.clearInterval;
    restorers.push(() => { window.setInterval = originalSetInterval; window.clearInterval = originalClearInterval; });
    return { active, started, cleared, advance(milliseconds: number) {
      const end = now + milliseconds;
      while (true) {
        const next = [...active.values()].reduce<number>((due, timer) => Math.min(due, timer.due), Infinity);
        if (next > end) break;
        now = next;
        for (const [id, timer] of [...active]) {
          if (!active.has(id) || timer.due !== next) continue;
          timer.due += timer.delay;
          timer.tick();
        }
      }
      now = end;
    } };
  }

  beforeEach(() => {
    timers = intervalClock();
    echoViews = [];
    echoRestores = 0;
    const visibility = Object.getOwnPropertyDescriptor(document, 'visibilityState');
    const dpr = Object.getOwnPropertyDescriptor(window, 'devicePixelRatio');
    Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' });
    Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 2 });
    restorers.push(() => {
      if (visibility) Object.defineProperty(document, 'visibilityState', visibility);
      else Reflect.deleteProperty(document, 'visibilityState');
      if (dpr) Object.defineProperty(window, 'devicePixelRatio', dpr);
      else Reflect.deleteProperty(window, 'devicePixelRatio');
      echoEnabled = false;
    });
    const echo = spyOn(peerlessPainterFactories, 'echo').mockImplementation((canvas, options) => {
      const painter = createLocalEchoPainter(canvas, options);
      return { ...painter, paintEcho(view) { echoViews.push(view); painter.paintEcho(view); },
        restoreLine(view) { echoRestores += 1; painter.restoreLine(view); } };
    });
    restorers.push(() => echo.mockRestore());
  });

  async function mountedCaret(configure?: (client: FakeCommandClient) => void) {
    const value = session();
    configure?.(value.client);
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} showToolbar={false} onReady={ready} />);
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const input = view.container.querySelector('textarea');
    expect(input === null).toBe(false);
    await act(async () => { input!.focus(); });
    return { ...value, view, input: input!, api: ready.mock.calls[0][0],
      canvas: view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement };
  }

  function textBoxFrame(frame: PresentationCommandFrame, update: (box: TextBoxPrimitive) => void) {
    const list = JSON.parse(new TextDecoder().decode(frame.displayList)) as SlideDisplayList;
    const box = list.primitives.find((primitive) => primitive.kind === 'textBox');
    if (!box || box.kind !== 'textBox') throw new Error('Expected text box');
    update(box);
    return { ...frame, displayList: new TextEncoder().encode(JSON.stringify(list)).buffer };
  }

  function caret(view: ReturnType<typeof render>): HTMLElement {
    const element = view.queryByTestId('pptx-peerless-caret');
    expect(element?.style.visibility).toBe('visible');
    return element!;
  }

  function paintingCounts() {
    return { frames: paints.length, echoes: echoViews.length, restores: echoRestores,
      canvases: [...canvasCalls.values()].map((entries) => entries.length) };
  }

  async function advance(milliseconds: number) {
    await act(async () => { timers.advance(milliseconds); });
  }

  async function visibility(state: 'visible' | 'hidden') {
    await act(async () => {
      Object.defineProperty(document, 'visibilityState', { configurable: true, value: state });
      fireEvent(document, new Event('visibilitychange'));
    });
  }

  test('blinks_caret_without_repainting_slide', async () => {
    const { view, input } = await mountedCaret();
    const element = caret(view);
    expect(element.getAttribute('aria-hidden')).toBe('true');
    expect(element.style.pointerEvents).toBe('none');
    expect(element.style.left).toBe('30px');
    expect(element.style.top).toBe('40px');
    expect(element.style.width).toBe('1px');
    expect(element.style.height).toBe('20px');
    expect(element.style.left).toBe(input.style.left);
    expect(element.style.top).toBe(input.style.top);
    expect(element.style.visibility).toBe('visible');
    expect([...timers.active.values()].map((timer) => timer.delay)).toEqual([530]);
    const counts = paintingCounts();
    await advance(530);
    expect(element.style.visibility).toBe('hidden');
    expect(paintingCounts()).toEqual(counts);
    await advance(530);
    expect(element.style.visibility).toBe('visible');
    expect(paintingCounts()).toEqual(counts);
  });

  test('restarts_caret_blink_on_movement_and_edits', async () => {
    const value = await mountedCaret();
    const element = caret(value.view);
    await advance(530);
    expect(element.style.visibility).toBe('hidden');
    await act(async () => { fireEvent.keyDown(value.input, { key: 'ArrowRight' }); });
    expect(element.style.left).toBe('40px');
    expect(element.style.visibility).toBe('visible');
    expect(timers.active.size).toBe(1);
    await advance(400);
    await act(async () => { fireEvent.keyDown(value.input, { key: 'ArrowLeft' }); });
    expect(element.style.left).toBe('30px');
    await advance(130);
    expect(element.style.visibility).toBe('visible');
    await advance(400);
    expect(element.style.visibility).toBe('hidden');
    await act(async () => { beforeInput(value.input, 'insertText', 'a'); });
    expect(element.style.visibility).toBe('visible');
    await advance(400);
    await act(async () => {
      value.client.ack(3, coveringFrame(value.client, value.client.text, 3));
    });
    expect(element.style.left).toBe('40px');
    await advance(130);
    expect(element.style.visibility).toBe('visible');
    await advance(400);
    expect(element.style.visibility).toBe('hidden');
    value.view.unmount();

    echoEnabled = true;
    const echoed = await mountedCaret();
    const echoCaret = caret(echoed.view);
    echoed.client.hold = true;
    await advance(530);
    await act(async () => { beforeInput(echoed.input, 'insertText', 'a'); });
    expect(echoCaret.style.left).toBe('40px');
    expect(echoCaret.style.visibility).toBe('visible');
    await act(async () => { echoed.client.ack(1, coveringFrame(echoed.client, 'aabc', 1)); });
    echoed.client.hold = false;
    await act(async () => { compose(echoed.input, 'compositionstart', ''); });
    await act(async () => { echoed.client.ack(2, coveringFrame(echoed.client, echoed.client.text, 2)); });
    await advance(530);
    await act(async () => { compose(echoed.input, 'compositionupdate', 'a'); });
    expect(echoCaret.style.left).toBe('50px');
    expect(echoCaret.style.visibility).toBe('visible');
    await advance(400);
    await act(async () => { compose(echoed.input, 'compositionupdate', 'ab'); });
    expect(echoCaret.style.left).toBe('60px');
    const compositionCounts = paintingCounts();
    await advance(130);
    expect(echoCaret.style.visibility).toBe('visible');
    await advance(400);
    expect(echoCaret.style.visibility).toBe('hidden');
    expect(paintingCounts()).toEqual(compositionCounts);
    expect(timers.active.size).toBe(1);
  });

  test('stops_caret_blinking_while_blurred_or_hidden', async () => {
    const value = await mountedCaret();
    expect(timers.active.size).toBe(1);
    const element = caret(value.view);
    const timer = timers.started[timers.started.length - 1];
    await act(async () => { value.input.blur(); });
    expect(element.style.visibility).toBe('hidden');
    expect(timers.cleared).toContain(timer);
    expect(timers.active.size).toBe(0);
    const blurred = paintingCounts();
    await advance(1060);
    expect(paintingCounts()).toEqual(blurred);
    await act(async () => { value.input.focus(); });
    expect(element.style.visibility).toBe('visible');
    expect(timers.active.size).toBe(1);
    await visibility('hidden');
    expect(element.style.visibility).toBe('hidden');
    expect(timers.active.size).toBe(0);
    const hidden = paintingCounts();
    await advance(1060);
    expect(paintingCounts()).toEqual(hidden);
    await visibility('visible');
    expect(element.style.visibility).toBe('visible');
    expect(timers.active.size).toBe(1);
    await advance(530);
    expect(element.style.visibility).toBe('hidden');
    await advance(530);
    expect(element.style.visibility).toBe('visible');
    await act(async () => { fireEvent.keyDown(value.input, { key: 'ArrowRight', shiftKey: true }); });
    expect(element.style.visibility).toBe('hidden');
    expect(timers.active.size).toBe(0);
    await act(async () => { await value.api.clearSelection(); });
    expect(element.style.visibility).toBe('hidden');
    expect(timers.active.size).toBe(0);
    await advance(1060);
    expect(element.style.visibility).toBe('hidden');
  });

  test('removes_canvas_caret_drawing', async () => {
    echoEnabled = true;
    const value = await mountedCaret();
    const caretStrokes = () => calls(value.canvas).filter((entry) => entry.kind === 'fillRect'
      && entry.args[2] === 1 && entry.args[3] === 20);
    expect(kinds(caretStrokes())).toEqual([]);
    calls(value.canvas).length = 0;
    await act(async () => { fireEvent.keyDown(value.input, { key: 'ArrowRight', shiftKey: true }); });
    expect(calls(value.canvas).some((entry) => entry.kind === 'fillRect'
      && entry.args.join(',') === '30,40,10,20')).toBe(true);
    await act(async () => { fireEvent.keyDown(value.input, { key: 'ArrowLeft' }); });
    await act(async () => { value.client.ack(2, coveringFrame(value.client, value.client.text, 2)); });
    value.client.hold = true;
    calls(value.canvas).length = 0;
    await act(async () => { beforeInput(value.input, 'insertText', 'a'); });
    expect(calls(value.canvas).filter((entry) => entry.kind === 'fillText')
      .map((entry) => entry.args[0]).join('')).toBe('aabc');
    expect(kinds(caretStrokes())).toEqual([]);
    const echo = echoViews[echoViews.length - 1];
    expect(echo.echo).not.toBeNull();
    const fallback = createLocalEchoPainter(value.canvas);
    try {
      calls(value.canvas).length = 0;
      fallback.paintEcho({ ...echo, selection: null });
      expect(kinds(caretStrokes())).toEqual([]);
    } finally { fallback.dispose(); }
    await act(async () => { value.client.ack(3, coveringFrame(value.client, 'aabc', 3)); });
    value.client.hold = false;
    await act(async () => { compose(value.input, 'compositionstart', ''); });
    await act(async () => { value.client.ack(4, coveringFrame(value.client, value.client.text, 4)); });
    calls(value.canvas).length = 0;
    await act(async () => { compose(value.input, 'compositionupdate', 'a'); });
    expect(calls(value.canvas).filter((entry) => entry.kind === 'fillText')
      .map((entry) => entry.args[0]).join('')).toBe('aaabc');
    expect(calls(value.canvas).some((entry) => entry.kind === 'fillRect'
      && entry.args.join(',') === '40,57,10,1')).toBe(true);
    expect(kinds(caretStrokes())).toEqual([]);
  });

  test('cleans_up_caret_blink_on_unmount', async () => {
    const value = await mountedCaret();
    expect(timers.active.size).toBe(1);
    const timer = timers.started[timers.started.length - 1];
    await act(async () => { value.view.unmount(); });
    expect(timers.cleared).toContain(timer);
    expect(timers.active.size).toBe(0);
    const counts = paintingCounts();
    await advance(1060);
    expect(paintingCounts()).toEqual(counts);
  });

  test('rotated_echo_and_ime_caret_is_a_thin_transformed_line', async () => {
    echoEnabled = true;
    const value = await mountedCaret((client) => {
      client.state.frame = textBoxFrame(client.state.frame, (box) => {
        box.transform = { rotationDeg: 30, flipH: true };
      });
    });
    const element = caret(value.view);
    const expectLine = (localX: number) => {
      const cos = Math.cos(Math.PI / 6);
      const sin = Math.sin(Math.PI / 6);
      expect(element.style.width).toBe('1px');
      expect(element.style.height).toBe('20px');
      expect(element.style.transformOrigin).toBe('0 0');
      expect(parseFloat(element.style.left)).toBeCloseTo(170 - (localX - 170) * cos + 20 * sin);
      expect(parseFloat(element.style.top)).toBeCloseTo(60 - (localX - 170) * sin - 20 * cos);
      const matrix = element.style.transform.match(/^matrix\(([^)]+)\)$/);
      expect(matrix).not.toBeNull();
      const values = matrix![1].split(',').map(Number);
      for (const [index, expected] of [-cos, -sin, -sin, cos, 0, 0].entries()) {
        expect(values[index]).toBeCloseTo(expected);
      }
      expect(element.style.visibility).toBe('visible');
    };
    expectLine(30);
    value.client.hold = true;
    await advance(530);
    await act(async () => { beforeInput(value.input, 'insertText', 'a'); });
    expect(echoViews[echoViews.length - 1].echo).not.toBeNull();
    expectLine(40);
    await act(async () => { value.client.ack(1, coveringFrame(value.client, 'aabc', 1)); });
    value.client.hold = false;
    await act(async () => { compose(value.input, 'compositionstart', ''); });
    await act(async () => { value.client.ack(2, coveringFrame(value.client, value.client.text, 2)); });
    await advance(530);
    await act(async () => { compose(value.input, 'compositionupdate', 'a'); });
    expect(echoViews[echoViews.length - 1].composition?.text).toBe('a');
    expectLine(50);
    const counts = paintingCounts();
    await advance(530);
    expect(element.style.visibility).toBe('hidden');
    expect(paintingCounts()).toEqual(counts);
  });

  test('clips_pending_echo_caret_to_active_line_and_containers', async () => {
    echoEnabled = true;
    for (const kind of [null, 'chart', 'table'] as const) {
      const value = await mountedCaret((client) => {
        client.state.frame = textBoxFrame(client.state.frame, (box) => {
          if (kind === null) {
            box.w = 35;
            box.transform = { rotationDeg: 90, flipH: true };
            box.textShadow = { color: '#000000', blur: 1, dx: 2, dy: 1 };
          }
        });
        if (kind !== null) {
          const list = JSON.parse(new TextDecoder().decode(client.state.frame.displayList)) as SlideDisplayList;
          list.primitives = [{ kind, objectId: 2, x: 20, y: 40, w: 45, h: 40,
            name: kind, label: kind, transform: { rotationDeg: 90, flipH: true }, primitives: list.primitives }];
          client.state.frame = { ...client.state.frame,
            displayList: new TextEncoder().encode(JSON.stringify(list)).buffer };
        }
      });
      const element = caret(value.view);
      const midpointIsClipped = () => {
        const matrix = element.style.transform.slice(7, -1).split(',').map(Number);
        const x = parseFloat(element.style.left) + matrix[0] * 0.5 + matrix[2] * 10;
        const y = parseFloat(element.style.top) + matrix[1] * 0.5 + matrix[3] * 10;
        expect(x > 0 && x < 960 && y > 0 && y < 540).toBe(true);
        for (let parent = element.parentElement; parent; parent = parent.parentElement) {
          if (!parent.style.clipPath.startsWith('polygon(')) continue;
          const points = parent.style.clipPath.slice(8, -1).split(',')
            .map((point) => point.trim().split(/\s+/).map(parseFloat));
          const crosses = points.map((point, index) => {
            const next = points[(index + 1) % points.length];
            return (next[0] - point[0]) * (y - point[1]) - (next[1] - point[1]) * (x - point[0]);
          });
          if (!crosses.every((cross) => cross >= -1e-6) && !crosses.every((cross) => cross <= 1e-6)) return true;
        }
        return false;
      };
      expect(midpointIsClipped()).toBe(false);
      value.client.hold = true;
      await act(async () => { beforeInput(value.input, 'insertText', 'aaaa'); });
      expect(echoViews[echoViews.length - 1].echo?.caretPosition).toBe(5);
      expect(element.style.visibility).toBe('visible');
      expect(midpointIsClipped()).toBe(true);
      await act(async () => { value.view.unmount(); });
    }
  });

  test('clips_off_slide_caret_at_right_and_bottom_edges', async () => {
    const value = await mountedCaret((client) => {
      client.state.frame = textBoxFrame(client.state.frame, (box) => {
        box.x = 950;
        box.y = 535;
        box.lines[0] = { ...box.lines[0], x: 950, y: 535, baseline: 550,
          caretStops: [950, 959, 961, 962].map((x, position) => ({ position, x })) };
      });
    });
    const element = caret(value.view);
    const clip = value.view.getByTestId('pptx-peerless-caret-clip');
    expect(element.parentElement === clip).toBe(true);
    expect(clip.style.position).toBe('absolute');
    expect(clip.style.left).toBe('0px');
    expect(clip.style.top).toBe('0px');
    expect(clip.style.width).toBe('960px');
    expect(clip.style.height).toBe('540px');
    expect(clip.style.overflow).toBe('hidden');
    expect(element.style.left).toBe('959px');
    expect(element.style.top).toBe('535px');
    expect(element.style.height).toBe('20px');
    await act(async () => { fireEvent.keyDown(value.input, { key: 'ArrowRight' }); });
    expect(element.style.left).toBe('961px');
    expect(element.parentElement === clip).toBe(true);
    await act(async () => {
      value.client.ack(1, textBoxFrame(coveringFrame(value.client, value.client.text, 1), (box) => {
        box.x = 20;
        box.y = 541;
        box.lines[0] = { ...box.lines[0], x: 20, y: 541, baseline: 556 };
      }));
    });
    expect(element.style.left).toBe('40px');
    expect(element.style.top).toBe('541px');
    expect(element.parentElement === clip).toBe(true);
    expect(clip.style.overflow).toBe('hidden');
    const counts = paintingCounts();
    await advance(530);
    expect(element.style.visibility).toBe('hidden');
    expect(paintingCounts()).toEqual(counts);
  });
});

describe('peerless stage fit', () => {
  let stageSize: { width: number; height: number };
  let beforeChromeWidth: number | undefined;
  const observers = new Map<ResizeObserver, { callback: ResizeObserverCallback; target: Element }>();

  function override(target: object, key: string, descriptor: PropertyDescriptor): void {
    const original = Object.getOwnPropertyDescriptor(target, key);
    Object.defineProperty(target, key, { configurable: true, ...descriptor });
    restorers.push(() => {
      if (original) Object.defineProperty(target, key, original);
      else Reflect.deleteProperty(target, key);
    });
  }

  function domScale(element: HTMLElement | null): number {
    let scale = 1;
    for (; element; element = element.parentElement) {
      const transform = element.style.transform.match(/^(?:scale(?:3d)?|matrix)\(\s*([-+\d.e]+)/i);
      scale *= (Number(transform?.[1]) || 1) * (Number.parseFloat(element.style.zoom) || 1);
    }
    return scale;
  }

  beforeEach(() => {
    stageSize = { width: 520, height: 310 };
    beforeChromeWidth = undefined;
    observers.clear();
    for (const dimension of ['width', 'height'] as const) {
      override(HTMLElement.prototype, dimension === 'width' ? 'clientWidth' : 'clientHeight', {
        get(this: HTMLElement) {
          if (this.getAttribute('role') !== 'application') return 0;
          if (dimension === 'width' && beforeChromeWidth !== undefined && !this.parentElement?.querySelector('aside')) {
            return beforeChromeWidth;
          }
          return stageSize[dimension];
        },
      });
    }
    override(window, 'devicePixelRatio', { value: 2 });
    const bounds = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
      if (this.getAttribute('role') === 'application') return new DOMRect(0, 0, this.clientWidth, this.clientHeight);
      const scale = domScale(this);
      return new DOMRect(0, 0, (Number.parseFloat(this.style.width) || 0) * scale,
        (Number.parseFloat(this.style.height) || 0) * scale);
    });
    restorers.push(() => bounds.mockRestore());
    class StageObserver {
      constructor(readonly callback: ResizeObserverCallback) {}
      observe(target: Element) { observers.set(this as unknown as ResizeObserver, { callback: this.callback, target }); }
      unobserve() { this.disconnect(); }
      disconnect() { observers.delete(this as unknown as ResizeObserver); }
    }
    override(globalThis, 'ResizeObserver', { value: StageObserver });
    override(window, 'ResizeObserver', { value: StageObserver });
    let nextIdle = 0;
    const cancelledIdle = new Set<number>();
    override(window, 'requestIdleCallback', { value: (callback: IdleRequestCallback) => {
      const id = ++nextIdle;
      queueMicrotask(() => {
        if (!cancelledIdle.has(id)) callback({ didTimeout: false, timeRemaining: () => 50 });
      });
      return id;
    } });
    override(window, 'cancelIdleCallback', { value: (id: number) => { cancelledIdle.add(id); } });
  });

  async function flushStage(action?: () => void): Promise<void> {
    await act(async () => {
      action?.();
      for (let index = 0; index < 60; index += 1) await Promise.resolve();
    });
    await act(async () => {
      for (let index = 0; index < 60; index += 1) await Promise.resolve();
    });
  }

  async function resize(width: number, height: number): Promise<void> {
    await flushStage(() => {
      stageSize = { width, height };
      for (const [observer, { callback, target }] of observers) {
        const contentRect = target.getBoundingClientRect();
        const size = [{ inlineSize: contentRect.width, blockSize: contentRect.height }];
        callback([{ target, contentRect, contentBoxSize: size, borderBoxSize: size,
          devicePixelContentBoxSize: size }], observer);
      }
    });
  }

  function heldBitmaps() {
    const requests: { gate: ReturnType<typeof deferred<ImageBitmap>>; source: ImageBitmap;
      close: ReturnType<typeof mock<() => void>> }[] = [];
    const decode = mock(() => {
      const gate = deferred<ImageBitmap>();
      const close = mock(() => {});
      const source = { width: 100, height: 100, close } as unknown as ImageBitmap;
      requests.push({ gate, source, close });
      return gate.promise;
    });
    override(globalThis, 'createImageBitmap', { value: decode });
    override(window, 'createImageBitmap', { value: decode });
    const releaseAll = () => { for (const { gate, source } of requests) gate.resolve(source); };
    restorers.push(releaseAll);
    return { requests, decode, releaseAll,
      release(index: number) { requests[index].gate.resolve(requests[index].source); } };
  }

  async function mountedStage() {
    const value = session(true);
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
    await flushStage(value.admit);
    expect(ready).toHaveBeenCalledTimes(1);
    return { ...value, view, ready, api: ready.mock.calls[0][0],
      canvas: view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement };
  }

  function drawnRaster(canvas: HTMLCanvasElement): HTMLCanvasElement {
    const draws = calls(canvas).filter((entry) => entry.kind === 'drawImage' && entry.args[0] instanceof HTMLCanvasElement);
    const source = draws[draws.length - 1]?.args[0];
    expect(source).toBeInstanceOf(HTMLCanvasElement);
    return source as HTMLCanvasElement;
  }

  function rasterTiles(canvas: HTMLCanvasElement): HTMLCanvasElement[] {
    return calls(canvas).filter((entry) => entry.kind === 'drawImage')
      .map((entry) => entry.args[0]).filter((source): source is HTMLCanvasElement => source instanceof HTMLCanvasElement);
  }

  test('scaled_overlays_stay_inside_committed_viewport', async () => {
    const value = await mountedStage();
    try {
      const input = value.view.container.querySelector('textarea')!;
      const viewport = input.parentElement!;
      const layer = value.view.getByTestId('pptx-peerless-caret-clip').parentElement!;
      expect(layer.parentElement).toBe(viewport);
      expect(value.view.getByTestId('pptx-peerless-caret-clip').style.overflow).toBe('hidden');
      expect(layer.style.width).toBe('960px');
      expect(Number.parseFloat(value.view.getByRole('application').style.minWidth)).toBe(0);
      for (const [width, height] of [[88, 100], [24, 14]]) {
        await resize(width, height);
        expect([viewport.style.width, viewport.style.height]).toEqual([value.canvas.style.width, value.canvas.style.height]);
        expect(Number.parseFloat(input.style.left)).toBeCloseTo(30 * Number.parseFloat(value.canvas.style.width) / 960);
        const since = calls(value.canvas).length;
        await resize(width, height);
        expect(calls(value.canvas).length).toBe(since);
      }
    } finally { value.view.unmount(); }
  });

  test('shared_decode_survives_creator_disposal', async () => {
    const images = heldBitmaps();
    const value = await mountedStage();
    const consumers: ReturnType<typeof createCommandFramePainter>[] = [];
    try {
      const raw = pictureFrame(value.client.state.frame, 'shared-pending');
      const frame = { frame: raw, displayList: pptx.decodeCommandDisplayList(raw) };
      const snapshot = { ...paintedViews[paintedViews.length - 1], confirmed: frame };
      const firstCanvas = document.createElement('canvas');
      const secondCanvas = document.createElement('canvas');
      const first = peerlessPainterFactories.frame(firstCanvas, { scale: 0.25 });
      const second = peerlessPainterFactories.frame(secondCanvas, { scale: 0.25 });
      consumers.push(first, second);
      await flushStage(() => { first.paintFrame(frame, snapshot); second.paintFrame(frame, snapshot); });
      const raster = commandFrameRaster(frame, 0.25)!;
      let painted = false;
      void second.whenPainted().then((result) => { painted = result; });
      first.dispose();
      expect(raster.canvas.width).toBeGreaterThan(0);
      await flushStage(() => { images.release(0); });
      expect(painted).toBe(true);
      expect(drawnRaster(secondCanvas)).toBe(raster.canvas);
      expect(rasterTiles(raster.canvas).some((tile) => calls(tile).some((entry) =>
        entry.kind === 'drawImage' && entry.args[0] === images.requests[0].source))).toBe(true);
      second.dispose();
      expect([raster.canvas.width, raster.canvas.height]).toEqual([0, 0]);
    } finally {
      for (const consumer of consumers) consumer.dispose();
      value.view.unmount();
      await flushStage(images.releaseAll);
    }
  });

  test('read_only_allows_selection_and_blocks_edits_until_mode_ack', async () => {
    const value = session(true);
    value.client.state.editState.readOnly = true;
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} readOnly onReady={ready} />);
    try {
      await flushStage(value.admit);
      expect(ready).toHaveBeenCalledTimes(1);
      const canvas = view.getByTestId('pptx-slide-canvas');
      const input = view.container.querySelector('textarea')!;
      const pointer = { pointerId: 1, isPrimary: true, button: 0, clientX: 20, clientY: 27.5 };
      await flushStage(() => { fireEvent.pointerDown(canvas, pointer); fireEvent.pointerUp(canvas, pointer); });
      expect(value.client.sent[0]).toMatchObject({ kind: 'selectAtPoint', payload: { x: 40, y: 55 } });
      expect(value.client.state.selection).toMatchObject({ anchor: 2, focus: 2 });
      await flushStage(() => { fireEvent.keyDown(input, { key: 'ArrowLeft' }); beforeInput(input, 'insertText', 'X'); });
      expect(value.client.sent.map((intent) => intent.kind)).toEqual(['selectAtPoint', 'setSelection']);
      expect(value.client.sent[1]).toMatchObject({ payload: { selection: { anchor: 1, focus: 1 } } });
      expect(value.client.text).toBe('abc');
      value.client.hold = true;
      await flushStage(() => { view.rerender(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />); });
      await flushStage(() => {
        fireEvent.pointerDown(canvas, pointer);
        fireEvent.pointerUp(canvas, pointer);
        beforeInput(input, 'insertText', 'X');
      });
      expect(value.client.sent.map((intent) => intent.kind)).toEqual(['selectAtPoint', 'setSelection', 'setReadOnly', 'selectAtPoint']);
      await flushStage(() => { value.client.ack(4); });
      value.client.hold = false;
      await flushStage(() => { beforeInput(input, 'insertText', 'X'); });
      expect(value.client.sent[4].kind).toBe('insertText');
    } finally { view.unmount(); }
  });

  test('opening_mode_ack_preserves_initial_media_before_fonts', async () => {
    const images = heldBitmaps();
    const loading = deferred<void>();
    const installed = new Set<FontFace>();
    class HeldFontFace {
      load() { return loading.promise.then(() => this as unknown as FontFace); }
    }
    override(globalThis, 'FontFace', { value: HeldFontFace });
    override(window, 'FontFace', { value: HeldFontFace });
    override(document, 'fonts', { value: { add(face: FontFace) { installed.add(face); },
      delete(face: FontFace) { return installed.delete(face); }, ready: Promise.resolve() } });
    const value = session(true);
    value.client.state.frame = pictureFrame(value.client.state.frame, 'opening-image');
    value.client.state.editState.readOnly = true;
    value.client.hold = true;
    const fonts = [{ family: 'Arial', bytes: new Uint8Array([1]) }];
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={fonts} file={file} {...{ [PEERLESS_FLAG]: true as const }} readOnly onReady={ready} />);
    try {
      await flushStage(() => { view.rerender(<PptxEditor fonts={fonts} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />); });
      await flushStage(value.admit);
      expect(value.client.sent[0]).toMatchObject({ kind: 'setReadOnly', payload: { readOnly: false } });
      await flushStage(() => {
        value.client.ack(1, { ...coveringFrame(value.client, 'abc', 1), media: [] });
      });
      expect(value.client.state.frame.media).toEqual([]);
      expect(paints).toEqual([]);
      expect(images.decode).not.toHaveBeenCalled();
      expect(ready).not.toHaveBeenCalled();
      await flushStage(() => { loading.resolve(); });
      expect(images.decode).toHaveBeenCalledTimes(1);
      expect(ready).not.toHaveBeenCalled();
      await flushStage(images.releaseAll);
      expect(ready).toHaveBeenCalledTimes(1);
      const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
      expect(rasterTiles(drawnRaster(canvas)).some((tile) => calls(tile).some((entry) =>
        entry.kind === 'drawImage' && entry.args[0] === images.requests[0].source))).toBe(true);
      view.unmount();
      expect(installed.size).toBe(0);
    } finally {
      view.unmount();
      await flushStage(() => { loading.resolve(); images.releaseAll(); });
    }
  });

  test('successive_frames_reuse_all_pinned_images_and_tiles', async () => {
    const images = heldBitmaps();
    const value = await mountedStage();
    try {
      const raw = pictureFrame(value.client.state.frame, 'image-0');
      const list = pptx.decodeCommandDisplayList(raw);
      const picture = list.primitives.pop()!;
      list.primitives.push(...Array.from({ length: 26 }, (_, index) => ({ ...picture, objectId: 10 + index,
        assetId: `image-${index}` })));
      raw.displayList = new TextEncoder().encode(JSON.stringify(list)).buffer;
      raw.media = Array.from({ length: 26 }, (_, index) => ({ assetId: `image-${index}`, bytes: new Uint8Array([1]).buffer }));
      await flushStage(() => { value.client.ack(0, raw); });
      expect(images.decode).toHaveBeenCalledTimes(26);
      await flushStage(images.releaseAll);
      const first = drawnRaster(value.canvas);
      const tiles = rasterTiles(first);
      const imageTiles = tiles.filter((tile) => calls(tile).some((entry) => entry.kind === 'drawImage' &&
        images.requests.some(({ source }) => source === entry.args[0])));
      expect(imageTiles).toHaveLength(26);
      const counts = imageTiles.map((tile) => calls(tile).length);
      await flushStage(() => { value.client.ack(0, coveringFrame(value.client, 'abc', 0)); });
      expect(drawnRaster(value.canvas)).not.toBe(first);
      expect(images.decode).toHaveBeenCalledTimes(26);
      expect(rasterTiles(drawnRaster(value.canvas)).every((tile, index) => tile === tiles[index])).toBe(true);
      expect(imageTiles.map((tile) => calls(tile).length)).toEqual(counts);
      const changed = coveringFrame(value.client, 'abc', 0);
      changed.media = [{ assetId: 'image-0', bytes: new Uint8Array([2]).buffer }];
      await flushStage(() => { value.client.ack(0, changed); });
      expect(images.decode).toHaveBeenCalledTimes(27);
      await flushStage(() => { images.release(26); });
      const updated = rasterTiles(drawnRaster(value.canvas));
      expect(imageTiles.filter((tile) => updated.includes(tile))).toHaveLength(25);
      await flushStage(() => { value.view.unmount(); });
      expect(images.requests.map(({ close }) => close.mock.calls.length)).toEqual(Array<number>(27).fill(1));
    } finally {
      value.view.unmount();
      await flushStage(images.releaseAll);
    }
  });

  test('strip_skips_caret_updates_and_renders_only_changed_preview', async () => {
    const client = new FakeCommandClient();
    client.state.slides.push({ id: 'other', index: 1, name: null, layoutPartPath: null });
    const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
    store.installThumbnail({ ...client.state.frame, slideId: 'other', slideIndex: 1 });
    const images = createCommandFrameImages();
    const run = { client, store, images, current: true, failure: undefined } as unknown as PeerlessPresentation;
    const previewFit = mock((_frame: PeerlessConfirmedFrame, _run: PeerlessPresentation) => ({ scale: 0.25, dpr: 2, generation: 1 }));
    const view = render(<PeerlessSlideStrip run={run} previewFit={previewFit} current={() => true} onNavigate={() => {}} />);
    try {
      await flushStage();
      previewFit.mockClear();
      await flushStage(() => { store.setFocused(true); store.setSelection({ ...selection, anchor: 2, focus: 2 }); });
      expect(previewFit).not.toHaveBeenCalled();
      await flushStage(() => { client.ack(0, coveringFrame(client, 'edited', 0)); });
      expect(previewFit.mock.calls.length).toBeGreaterThan(0);
      expect(previewFit.mock.calls.every(([frame]) => frame.frame.slideId === 'slide')).toBe(true);
      expect(view.container.querySelectorAll('canvas')).toHaveLength(2);
    } finally { view.unmount(); store.dispose(); images.dispose(); }
  });

  test('strip_skips_parent_rerenders', async () => {
    const value = session(true);
    const name = mock(() => null);
    const slide = { id: 'other', index: 1, name: null, layoutPartPath: null };
    Object.defineProperty(slide, 'name', { configurable: true, get: name });
    value.client.state.slides.push(slide);
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} />);
    try {
      await flushStage(value.admit);
      expect(name).toHaveBeenCalled();
      name.mockClear();
      await flushStage(() => {
        view.rerender(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} className="changed" />);
      });
      expect(name).not.toHaveBeenCalled();
      expect(view.container.querySelectorAll('aside button')).toHaveLength(2);
    } finally { view.unmount(); }
  });

  test('superseded_thumbnail_caches_media_and_repaints_referencing_frames', async () => {
    const images = heldBitmaps();
    const value = session(true);
    const first = deferred<PresentationCommandThumbnailResult>();
    const picture = pictureFrame(value.client.state.frame, 'shared-image');
    value.client.state.frame = { ...picture, media: [] };
    value.client.state.slides.push({ id: 'cached', index: 1, name: null, layoutPartPath: null },
      { id: 'other', index: 2, name: null, layoutPartPath: null });
    value.client.thumbnail.mockImplementation((slideId) => {
      if (slideId === 'cached') return Promise.resolve({ session: value.client.state.session,
        incarnation: value.client.state.incarnation, epoch: value.client.state.epoch, slideId, slideRevision: 0,
        status: 'frame', frame: { ...picture, slideId, slideIndex: 1, media: [] } });
      if (value.client.thumbnail.mock.calls.length === 2) return first.promise;
      return Promise.resolve({ session: value.client.state.session, incarnation: value.client.state.incarnation,
        epoch: value.client.state.epoch, slideId, slideRevision: 1, status: 'frame',
        frame: { ...picture, slideId, slideIndex: 2, slideRevision: 1, media: [] } });
    });
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} />);
    try {
      await flushStage(value.admit);
      expect(value.client.thumbnail).toHaveBeenCalledTimes(2);
      await flushStage(() => {
        for (const listener of value.client.listeners) listener({ ...value.client.state,
          frame: undefined, dirtySlideIds: ['other'] });
        first.resolve({ session: value.client.state.session, incarnation: value.client.state.incarnation,
          epoch: value.client.state.epoch, slideId: 'other', slideRevision: 0, status: 'frame',
          frame: { ...picture, slideId: 'other', slideIndex: 2 } });
      });
      expect(images.decode).toHaveBeenCalledTimes(1);
      await flushStage(images.releaseAll);
      await flushStage(images.releaseAll);
      expect(value.client.thumbnail).toHaveBeenCalledTimes(3);
      expect(view.container.querySelectorAll('aside canvas')).toHaveLength(3);
      const canvases = [view.getByTestId('pptx-slide-canvas'), ...view.container.querySelectorAll('aside canvas')];
      for (const canvas of canvases) {
        expect(rasterTiles(drawnRaster(canvas as HTMLCanvasElement)).some((tile) => calls(tile).some((entry) =>
          entry.kind === 'drawImage' && entry.args[0] === images.requests[0].source))).toBe(true);
      }
    } finally {
      view.unmount();
      first.resolve({ session: value.client.state.session, incarnation: value.client.state.incarnation,
        epoch: value.client.state.epoch, slideId: 'other', slideRevision: 0, status: 'unchanged' });
      await flushStage(images.releaseAll);
    }
  });

  test('stage_refits_retain_unchanged_thumbnail_rasters', async () => {
    const value = session(true);
    value.client.state.slides.push({ id: 'other', index: 1, name: null, layoutPartPath: null });
    value.client.thumbnail.mockImplementation((slideId) => Promise.resolve({ session: value.client.state.session,
      incarnation: value.client.state.incarnation, epoch: value.client.state.epoch, slideId, slideRevision: 0,
      status: 'frame', frame: { ...value.client.state.frame, slideId, slideIndex: 1 } }));
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} />);
    try {
      await flushStage(value.admit);
      const canvas = view.container.querySelector('aside canvas') as HTMLCanvasElement;
      const raster = drawnRaster(canvas);
      const since = calls(canvas).length;
      const count = paints.filter((entry) => entry.frame.slideId === 'other').length;
      for (const [width, height] of [[280, 175], [0, 0], [160, 107.5]]) await resize(width, height);
      expect(view.container.querySelector('aside canvas')).toBe(canvas);
      expect(drawnRaster(canvas)).toBe(raster);
      expect(calls(canvas).length).toBe(since);
      expect(paints.filter((entry) => entry.frame.slideId === 'other')).toHaveLength(count);
      expect(raster.width).toBeGreaterThan(0);
    } finally { view.unmount(); }
  });

  test('overlapping_resizes_release_superseded_rasters', async () => {
    const images = heldBitmaps();
    const value = await mountedStage();
    try {
      const displayed = drawnRaster(value.canvas);
      const geometry = [value.canvas.style.width, value.canvas.style.height, value.canvas.width, value.canvas.height];
      const since = calls(value.canvas).length;
      const beforeA = new Set(canvasCalls.keys());
      const frame = pictureFrame(coveringFrame(value.client, 'abc', 0), 'image-0');
      const list = pptx.decodeCommandDisplayList(frame);
      const picture = list.primitives.pop()!;
      list.primitives.push(...Array.from({ length: 26 }, (_, index) => ({ ...picture, objectId: 10 + index,
        assetId: `image-${index}` })));
      frame.displayList = new TextEncoder().encode(JSON.stringify(list)).buffer;
      frame.media = Array.from({ length: 26 }, (_, index) => ({ assetId: `image-${index}`, bytes: new Uint8Array([1]).buffer }));
      await flushStage(() => { value.client.ack(0, frame); });
      expect(images.decode).toHaveBeenCalledTimes(26);
      const abandonedA = [...canvasCalls.keys()].filter((canvas) => !beforeA.has(canvas) && !canvas.isConnected);
      const pendingRasters = (scale: number) => [...new Set(paints.map((entry) => commandFrameRaster(entry, scale))
        .filter((raster): raster is NonNullable<typeof raster> => raster !== undefined && !raster.ready))];
      const pendingA = pendingRasters(0.5);
      expect(pendingA).toHaveLength(1);
      const beforeB = new Set(canvasCalls.keys());
      await resize(280, 175);
      expect(pendingA.map(({ canvas }) => [canvas.width, canvas.height])).toEqual([[0, 0]]);
      const abandonedB = [...canvasCalls.keys()].filter((canvas) => !beforeB.has(canvas) && !canvas.isConnected);
      const pendingB = pendingRasters(0.25);
      expect(pendingB).toHaveLength(1);
      await resize(160, 107.5);
      expect(pendingB.map(({ canvas }) => [canvas.width, canvas.height])).toEqual([[0, 0]]);
      expect(images.decode).toHaveBeenCalledTimes(26);
      expectIntact(value.canvas, displayed, since);
      expect([value.canvas.style.width, value.canvas.style.height, value.canvas.width, value.canvas.height]).toEqual(geometry);
      let hydrated = false;
      void value.api.whenHydrated().then(() => { hydrated = true; }).catch(() => {});
      await flushStage(images.releaseAll);
      expect(hydrated).toBe(true);
      expect(value.canvas.style.width).toBe('120px');
      expect(value.canvas.style.height).toBe('67.5px');
      expect([value.canvas.width, value.canvas.height]).toEqual([240, 135]);
      expect(pendingRasters(128 / 960)).toHaveLength(0);
      await resize(280, 175);
      expect(value.canvas.style.width).toBe('240px');
      expect(images.decode).toHaveBeenCalledTimes(26);
      expect(images.requests.every(({ close }) => close.mock.calls.length === 0)).toBe(true);
      await flushStage(() => { value.view.unmount(); });
      expect([...abandonedA, ...abandonedB].every((canvas) => canvas.width === 0 && canvas.height === 0)).toBe(true);
      expect(images.requests.map(({ close }) => close.mock.calls.length)).toEqual(Array<number>(26).fill(1));
    } finally {
      value.view.unmount();
      await flushStage(images.releaseAll);
    }
  });

  test('first_frame_fits_post_chrome_48px', async () => {
    stageSize = { width: 88, height: 100 };
    beforeChromeWidth = 112;
    const value = session(true);
    const published: (string | number)[][] = [];
    const ready = mock((_api: PptxAsyncEditorApi) => {
      published.push([canvas.style.width, canvas.style.height, canvas.width, canvas.height]);
    });
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
    const canvas = await view.findByTestId('pptx-slide-canvas') as HTMLCanvasElement;
    try {
      await flushStage(value.admit);
      expect(published).toEqual([['48px', '27px', 96, 54]]);
      expect(view.container.querySelector('aside')).not.toBeNull();
      expect([...observers.values()].filter(({ target }) => target === view.getByRole('application'))).toHaveLength(1);
      const widths = calls(canvas).filter((entry) => entry.kind === 'resize' && entry.args[0] === 'width' && Number(entry.args[1]) > 0)
        .map((entry) => entry.args[1]);
      expect(widths.length).toBeGreaterThan(0);
      expect(widths.every((width) => width === 96)).toBe(true);
      const input = view.container.querySelector('textarea')!;
      await flushStage(() => { input.focus(); });
      await resize(24, 14);
      expect([canvas.style.width, canvas.style.height, canvas.width, canvas.height]).toEqual(['24px', '13.5px', 48, 27]);
      const caret = view.getByTestId('pptx-peerless-caret');
      expect(Number.parseFloat(input.style.left) * domScale(input.parentElement)).toBeCloseTo(0.75);
      expect(Number.parseFloat(input.style.top) * domScale(input.parentElement)).toBeCloseTo(1);
      expect(Number.parseFloat(caret.style.left) * domScale(caret.parentElement)).toBeCloseTo(0.75);
      expect(Number.parseFloat(caret.style.top) * domScale(caret.parentElement)).toBeCloseTo(1);
      expect(caret.style.visibility).toBe('visible');
      expect(await ready.mock.calls[0][0].getPositionAtPoint(0.75, 1.25)).toMatchObject({ kind: 'text', position: 1 });
      const pointer = { pointerId: 1, isPrimary: true, button: 0, clientX: 1, clientY: 1.375 };
      await flushStage(() => { fireEvent.pointerDown(canvas, pointer); fireEvent.pointerUp(canvas, pointer); });
      expect(value.client.sent[0]).toMatchObject({ kind: 'selectAtPoint', payload: { x: 40, y: 55 } });
      expect(value.client.state.selection).toMatchObject({ anchor: 2, focus: 2 });
      const since = calls(canvas).length;
      await resize(0, 0);
      expect(calls(canvas).slice(since).some((entry) => ['resize', 'clearRect', 'drawImage'].includes(entry.kind))).toBe(false);
      expect(ready).toHaveBeenCalledTimes(1);
    } finally { view.unmount(); }
  });

  test('file_replacement_releases_old_rasters_and_fits_first_frame', async () => {
    const images = heldBitmaps();
    const value = session(true);
    value.client.state.frame = pictureFrame(value.client.state.frame, 'old-stage');
    value.client.state.slides.push({ id: 'other', index: 1, name: null, layoutPartPath: null });
    value.client.thumbnail.mockImplementation((slideId) => Promise.resolve({ session: value.client.state.session,
      incarnation: value.client.state.incarnation, epoch: value.client.state.epoch, slideId, slideRevision: 0,
      status: 'frame', frame: { ...pictureFrame(value.client.state.frame, 'old-thumbnail'), slideId, slideIndex: 1 } }));
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
    const canvas = await view.findByTestId('pptx-slide-canvas') as HTMLCanvasElement;
    const oldRasters: { frame: PeerlessConfirmedFrame; scale: number;
      raster: NonNullable<ReturnType<typeof commandFrameRaster>>; tiles: HTMLCanvasElement[] }[] = [];
    const collect = () => {
      for (const frame of new Set(paints)) for (const scale of [0.5, 0.25, 128 / 960]) {
        const raster = commandFrameRaster(frame, scale);
        if (raster && !oldRasters.some((entry) => entry.raster === raster)) {
          oldRasters.push({ frame, scale, raster, tiles: rasterTiles(raster.canvas) });
        }
      }
    };
    try {
      await flushStage(value.admit);
      await flushStage(images.releaseAll);
      await flushStage(images.releaseAll);
      expect(ready).toHaveBeenCalledTimes(1);
      expect(view.container.querySelector('aside canvas')).not.toBeNull();
      expect(images.decode).toHaveBeenCalledTimes(2);
      collect();
      expect(oldRasters).toHaveLength(3);
      await flushStage(() => { value.client.ack(0, pictureFrame(coveringFrame(value.client, 'abc', 0), 'old-pending')); });
      expect(images.decode).toHaveBeenCalledTimes(3);
      collect();
      await resize(280, 175);
      collect();
      expect(oldRasters.filter(({ scale }) => scale === 0.25)).toHaveLength(1);
      expect(oldRasters.some(({ scale, raster }) => scale === 0.25 && !raster.ready && raster.canvas.width > 0)).toBe(true);
      const replacement = session(true);
      const replacementFile = new Uint8Array([2]);
      const since = calls(canvas).length;
      await flushStage(() => {
        stageSize = { width: 88, height: 100 };
        beforeChromeWidth = 112;
        view.rerender(<PptxEditor fonts={[]} file={replacementFile} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
      });
      expect(value.owner.dispose).toHaveBeenCalledTimes(1);
      expect(view.container.querySelector('aside')).toBeNull();
      for (const { frame, scale, raster, tiles } of oldRasters) {
        expect(commandFrameRaster(frame, scale)).toBeUndefined();
        expect([raster.canvas.width, raster.canvas.height]).toEqual([0, 0]);
        expect(tiles.map((tile) => [tile.width, tile.height])).toEqual(tiles.map(() => [0, 0]));
      }
      expect(images.requests.slice(0, 2).map(({ close }) => close.mock.calls.length)).toEqual([1, 1]);
      await flushStage(replacement.admit);
      expect(ready).toHaveBeenCalledTimes(2);
      expect([canvas.style.width, canvas.style.height, canvas.width, canvas.height]).toEqual(['48px', '27px', 96, 54]);
      const frame = paints.find((entry) => entry.frame === replacement.client.state.frame)!;
      const raster = commandFrameRaster(frame, 48 / 960)!;
      expect(drawnRaster(canvas)).toBe(raster.canvas);
      expect(calls(canvas).slice(since).filter((entry) => entry.kind === 'drawImage').map((entry) => entry.args[0]))
        .toEqual([raster.canvas]);
      expect(calls(canvas).slice(since).filter((entry) => entry.kind === 'resize' && entry.args[0] === 'width')
        .map((entry) => entry.args[1])).toEqual([96]);
      expect(view.container.querySelectorAll('aside')).toHaveLength(1);
      const committedSince = calls(canvas).length;
      await flushStage(images.releaseAll);
      expectIntact(canvas, raster.canvas, committedSince);
      expect(images.requests[2].close).toHaveBeenCalledTimes(1);
    } finally {
      view.unmount();
      await flushStage(images.releaseAll);
    }
  });

  test('stale_scale_paint_is_dropped', async () => {
    const images = heldBitmaps();
    const value = await mountedStage();
    try {
      const original = value.client.state.frame;
      const beforeA = new Set(canvasCalls.keys());
      await flushStage(() => { value.client.ack(0, pictureFrame(coveringFrame(value.client, 'abc', 0, original), 'slow')); });
      const abandoned = [...canvasCalls.keys()].filter((canvas) => !beforeA.has(canvas) && !canvas.isConnected);
      let hydrated = false;
      void value.api.whenHydrated().then(() => { hydrated = true; }).catch(() => {});
      await resize(280, 175);
      await flushStage(() => { value.client.ack(0, pictureFrame(coveringFrame(value.client, 'abc', 0, original), 'fast')); });
      expect(images.decode).toHaveBeenCalledTimes(2);
      await flushStage(() => { images.release(1); });
      expect(hydrated).toBe(true);
      expect([value.canvas.style.width, value.canvas.style.height, value.canvas.width, value.canvas.height])
        .toEqual(['240px', '135px', 480, 270]);
      expect(abandoned.map((canvas) => [canvas.width, canvas.height])).toEqual(abandoned.map(() => [0, 0]));
      const committed = drawnRaster(value.canvas);
      const since = calls(value.canvas).length;
      await flushStage(() => { images.release(0); });
      expectIntact(value.canvas, committed, since);
      expect(abandoned.every((canvas) => canvas.width === 0 && canvas.height === 0)).toBe(true);
      expect(value.ready).toHaveBeenCalledTimes(1);
    } finally {
      value.view.unmount();
      await flushStage(images.releaseAll);
    }
  });

  test('shared_raster_releases_only_after_last_consumer', async () => {
    const value = await mountedStage();
    const consumers: ReturnType<typeof createCommandFramePainter>[] = [];
    try {
      const raster = drawnRaster(value.canvas);
      const frame = paints.find((entry) => entry.frame === value.client.state.frame)!;
      const snapshot = paintedViews.find((entry) => entry.confirmed === frame)!;
      const scale = Number.parseFloat(value.canvas.style.width) / frame.displayList.width;
      const secondCanvas = document.createElement('canvas');
      const second = peerlessPainterFactories.frame(secondCanvas, { scale });
      consumers.push(second);
      await flushStage(() => { second.paintFrame(frame, snapshot); });
      expect(drawnRaster(secondCanvas)).toBe(raster);
      second.dispose();
      expect(raster.width).toBeGreaterThan(0);
      expect(raster.height).toBeGreaterThan(0);
      const thirdCanvas = document.createElement('canvas');
      const third = peerlessPainterFactories.frame(thirdCanvas, { scale });
      consumers.push(third);
      await flushStage(() => { third.paintFrame(frame, snapshot); });
      expect(drawnRaster(thirdCanvas)).toBe(raster);
      const tiles = [...new Set(calls(raster).filter((entry) => entry.kind === 'drawImage')
        .map((entry) => entry.args[0]).filter((source): source is HTMLCanvasElement => source instanceof HTMLCanvasElement))];
      await flushStage(() => { value.view.unmount(); });
      expect(raster.width).toBeGreaterThan(0);
      expect(tiles.every((tile) => tile.width > 0 && tile.height > 0)).toBe(true);
      third.dispose();
      expect([raster.width, raster.height]).toEqual([0, 0]);
      expect(tiles.map((tile) => [tile.width, tile.height])).toEqual(tiles.map(() => [0, 0]));
      for (const canvas of [raster, ...tiles]) {
        expect(calls(canvas).filter((entry) => entry.kind === 'resize' && entry.args[0] === 'width' && entry.args[1] === 0))
          .toHaveLength(1);
      }
    } finally {
      value.view.unmount();
      for (const consumer of consumers) consumer.dispose();
    }
  });

  test('font_gate_covers_constructor_preview_and_refit', async () => {
    stageSize = { width: 88, height: 100 };
    beforeChromeWidth = 112;
    const loading = deferred<void>();
    const installed = new Set<FontFace>();
    class HeldFontFace {
      load() { return loading.promise.then(() => this as unknown as FontFace); }
    }
    override(globalThis, 'FontFace', { value: HeldFontFace });
    override(window, 'FontFace', { value: HeldFontFace });
    override(document, 'fonts', { value: { add(face: FontFace) { installed.add(face); },
      delete(face: FontFace) { return installed.delete(face); }, ready: Promise.resolve() } });
    const value = session(true);
    value.client.state.slides.push({ id: 'other', index: 1, name: null, layoutPartPath: null });
    value.client.thumbnail.mockImplementation((slideId) => Promise.resolve({ session: value.client.state.session,
      incarnation: value.client.state.incarnation, epoch: value.client.state.epoch, slideId, slideRevision: 0,
      status: 'frame', frame: { ...value.client.state.frame, slideId, slideIndex: 1 } }));
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const fonts = [{ family: 'Arial', bytes: new Uint8Array([1]) }];
    const view = render(<PptxEditor fonts={fonts} file={file} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
    try {
      await flushStage(value.admit);
      await resize(96, 100);
      await resize(88, 100);
      await flushStage(() => { value.client.ack(0, coveringFrame(value.client, 'abc', 0)); });
      expect(paints).toEqual([]);
      expect([...canvasCalls.values()].flat().filter((entry) =>
        ['resize', 'clearRect', 'drawImage', 'fillText', 'fillRect'].includes(entry.kind))).toEqual([]);
      expect(view.container.querySelector('aside canvas')).toBeNull();
      expect(value.client.thumbnail).not.toHaveBeenCalled();
      expect(ready).not.toHaveBeenCalled();
      await flushStage(() => { loading.resolve(); });
      const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
      expect([canvas.style.width, canvas.style.height, canvas.width, canvas.height]).toEqual(['48px', '27px', 96, 54]);
      expect(ready).toHaveBeenCalledTimes(1);
      expect(installed.size).toBe(1);
      expect(view.container.querySelector('aside canvas')).not.toBeNull();
      expect(paints.some((frame) => frame.frame.slideId === 'other')).toBe(true);
      await resize(64, 53.5);
      expect(canvas.style.width).toBe('24px');
      await flushStage(() => { view.unmount(); });
      expect(installed.size).toBe(0);
    } finally {
      view.unmount();
      await flushStage(() => { loading.resolve(); });
    }
  });

  test('pending_decode_mode_ack_gates_input_and_ready', async () => {
    const images = heldBitmaps();
    const value = session(true);
    value.client.state = { ...value.client.state, frame: pictureFrame(value.client.state.frame, 'mode-image'),
      editState: { ...value.client.state.editState, readOnly: true } };
    value.client.hold = true;
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} readOnly onReady={ready} />);
    try {
      await flushStage(value.admit);
      for (const readOnly of [false, true, false]) {
        await flushStage(() => { view.rerender(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }}
          readOnly={readOnly} onReady={ready} />); });
      }
      expect(value.client.sent.map((intent) => ({ kind: intent.kind, payload: intent.payload }))).toMatchObject([
        { kind: 'setReadOnly', payload: { readOnly: false } },
        { kind: 'setReadOnly', payload: { readOnly: true } },
        { kind: 'setReadOnly', payload: { readOnly: false } },
      ]);
      expect(images.decode).toHaveBeenCalledTimes(1);
      await flushStage(() => { images.release(0); });
      expect(ready).not.toHaveBeenCalled();
      await flushStage(() => { value.client.ack(1); });
      await flushStage(() => {
        const input = view.container.querySelector('textarea');
        if (input) beforeInput(input, 'insertText', 'X');
      });
      expect(value.client.sent.map((intent) => intent.kind)).toEqual(['setReadOnly', 'setReadOnly', 'setReadOnly']);
      expect(ready).not.toHaveBeenCalled();
      await flushStage(() => { value.client.ack(2); });
      expect(ready).not.toHaveBeenCalled();
      await flushStage(() => { value.client.ack(3); });
      expect(ready).toHaveBeenCalledTimes(1);
      let hydrated = false;
      void ready.mock.calls[0][0].whenHydrated().then(() => { hydrated = true; }).catch(() => {});
      await flushStage();
      expect(hydrated).toBe(true);
      value.client.hold = false;
      await flushStage(() => { beforeInput(view.container.querySelector('textarea')!, 'insertText', 'X'); });
      expect(value.client.sent[3]).toMatchObject({ kind: 'insertText', payload: { text: 'X' } });
      expect(value.client.text).toBe('aXbc');
    } finally {
      view.unmount();
      await flushStage(images.releaseAll);
    }
  });
});

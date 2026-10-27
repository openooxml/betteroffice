import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import {
  COMMAND_OUTCOME_UNKNOWN,
  isCommandOutcomeUnknownError,
  type Checkpoint,
  type CommandConnection,
  type EchoHooks,
} from '../../../../shared/office-session/commands/types';
import { isTransferResult, transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import { SessionFailure } from '../../../../shared/office-session/types';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { createDocxWorkerEditSession } from './createDocxWorkerEditSession';
import { createDocxCommandConnection } from './residentCommandConnection';
import { buildCommandFrame } from './residentCommandFrame';
import { createResidentCommandHost, packCommandAck } from './residentCommandHost';
import type {
  DocxCommandAck,
  DocxCommandApplyContext,
  DocxCommandEnvelope,
  DocxCommandFrame,
  DocxCommandFramePayload,
  DocxCommandPayload,
  DocxCommandSelection,
  DocxEditState,
  DocxWorkerEditSession,
  DocxWorkerEditSessionOptions,
  DocxCommandClient,
  DocxCommandConnection,
  DocxCommandHooks,
  DocxCommandHost,
  DocxCommandHostOptions,
  DocxCommandInitRequest,
  DocxCommandInitReply,
  DocxCommandRegistry,
} from './residentCommandProtocol';
import type { ResidentEngineSession } from './residentEngineSession';

const TIMEOUT = 15_000;
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
const INIT: DocxCommandInitRequest = {
  session: 'docx-command-test', incarnation: 3, epoch: 7, readOnly: false,
  fonts: [], layoutInput: LAYOUT, layoutWithRegions: true, layoutRevision: 1, extras: '{}',
};

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function unwrap<T>(reply: T | TransferResult<T>): T {
  return isTransferResult(reply) ? (reply as TransferResult<T>).value : reply;
}

function createHost(options: DocxCommandHostOptions): DocxCommandHost {
  let host!: DocxCommandHost;
  expect(() => { host = createResidentCommandHost(options); }).not.toThrow();
  return host;
}

function createSession<X>(options: DocxWorkerEditSessionOptions<Uint8Array, X>): DocxWorkerEditSession<X> {
  let session!: DocxWorkerEditSession<X>;
  expect(() => { session = createDocxWorkerEditSession(options); }).not.toThrow();
  return session;
}

function envelope(seq: number, extra: Partial<DocxCommandEnvelope> = {}): DocxCommandEnvelope {
  return {
    session: INIT.session, incarnation: INIT.incarnation, seq,
    basis: { epoch: INIT.epoch, version: 0 }, kind: 'insertText', payload: { text: 'a' },
    ...extra,
  };
}

function emptyAck(extra: Partial<DocxCommandAck> = {}): DocxCommandAck {
  return {
    session: INIT.session, incarnation: INIT.incarnation, epoch: INIT.epoch,
    version: 0, throughSeq: 0, outcomes: [], ...extra,
  };
}

function initialAck(extra: Partial<DocxCommandInitReply> = {}): DocxCommandInitReply {
  return {
    ...emptyAck(), selection: null,
    editState: {
      readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null,
    },
    frame: {
      coversThroughSeq: 0, revision: 0,
      payload: { frame: new ArrayBuffer(0), caret: { frameEpoch: 0, caretRect: null }, layoutRevision: 1 },
    },
    ...extra,
  };
}

function machine() {
  const point = { story: 'body', paraId: 'p1', offset: 0 };
  const selection: DocxCommandSelection = { anchor: { ...point }, focus: { ...point } };
  const editState: DocxEditState = {
    readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
    canUndo: false, canRedo: false, hyperlink: null, tableContext: null,
  };
  const state = {
    text: '', selection, editState,
    snapshot: new Uint8Array([11, 22, 33]),
    contexts: [] as DocxCommandApplyContext[],
    applied: [] as number[],
    preflighted: [] as number[],
  };
  let revision = 40;
  const hooks: DocxCommandHooks = {
    preflight(command, context) {
      state.contexts.push(context);
      state.preflighted.push(command.seq);
      if (command.kind !== 'insertText' || !('text' in command.payload)) {
        return { status: 'refused', reason: 'invalid-payload' };
      }
      if (command.payload.text === 'refuse') return { status: 'refused', reason: 'locked' };
      return {
        status: 'ready',
        prepared: {
          command: { ...command, kind: 'insertText', payload: { text: command.payload.text } },
          selection: state.selection, range: null,
        },
      };
    },
    apply({ command }) {
      if (command.kind !== 'insertText') throw new Error('unexpected test command');
      if (command.payload.text === '') return { status: 'noop' };
      state.applied.push(command.seq);
      state.text += command.payload.text;
      return { status: 'applied' };
    },
    selection: () => state.selection,
    editState: () => state.editState,
    frame({ coversThroughSeq }) {
      return {
        coversThroughSeq, revision: ++revision,
        payload: { frame: new ArrayBuffer(0), caret: { frameEpoch: revision, caretRect: null }, layoutRevision: 1 },
      };
    },
    snapshot: () => state.snapshot,
  };
  const registry: DocxCommandRegistry = { handlers: {}, policies: {} };
  return { state, hooks, registry, options: { init: INIT, hooks, registry } satisfies DocxCommandHostOptions };
}

function caretEditState(paraId: string, bold = false): DocxEditState {
  return {
    ...initialAck().editState, documentVersion: 'unchanged',
    toolbar: {
      paraId, bold, italic: false, underline: false, strike: false, superscript: false, subscript: false,
      fontFamily: 'Arial', fontSize: 24, color: null, highlight: null, styleId: null, alignment: null,
      paragraphProperties: {}, hasSelection: false, isMultiParagraph: false, inTable: false,
      isSingleEmbed: false, embedKind: null, isImage: false, inInsertion: false, inDeletion: false,
    },
    effectiveFormatting: {
      bold, italic: false, underline: false, strike: false, superscript: false, subscript: false,
      fontFamily: 'Arial', fontSizePt: 12, color: null, highlight: null, styleId: null, alignment: null,
    },
  };
}

test('selection noop acknowledgements defer frames only with selection-derived edit state changes and worker opt-in', async () => {
  for (const mode of ['selection', 'cross-paragraph', 'no-context', 'stored-marks', 'edit-state', 'document-version', 'undo', 'redo', 'content', 'non-selection', 'opt-out'] as const) {
    const m = machine();
    if (mode === 'cross-paragraph') m.state.editState = caretEditState('p1');
    if (mode === 'stored-marks') m.state.editState.storedMarks = { story: 'body', paraId: 'p1', clear: false, delta: { bold: true } };
    const deferredAcks: DocxCommandAck[] = [];
    const echoContext = { page: { pageId: '1', pageIndex: 0 }, frameEpoch: 40, layoutRevision: 1,
      serializedBytes: 100, text: 'ab', styleRuns: [], boundaries: [],
      selection: { anchorGrapheme: 1, focusGrapheme: 1 }, transform: [1, 0, 0, 1, 0, 0],
      damage: [], shaping: 'simple-ltr' } satisfies NonNullable<DocxCommandAck['echoContext']>;
    m.hooks.preflight = command => ({ status: 'ready', prepared: {
      command: { ...command, kind: 'select', payload: { selection: m.state.selection } },
      selection: m.state.selection, range: null,
    } });
    m.hooks.apply = () => {
      m.state.selection = { ...m.state.selection, focus: { ...m.state.selection.focus, offset: 1 } };
      if (mode === 'cross-paragraph') {
        const at = { story: 'body', paraId: 'p2', offset: 1 };
        m.state.selection = { anchor: at, focus: at };
        m.state.editState = caretEditState('p2', true);
      }
      if (mode === 'stored-marks') m.state.editState.storedMarks = null;
      if (mode === 'edit-state') m.state.editState.readOnly = true;
      if (mode === 'document-version') m.state.editState.documentVersion = 'changed';
      if (mode === 'undo') m.state.editState.canUndo = true;
      if (mode === 'redo') m.state.editState.canRedo = true;
      return { status: mode === 'content' ? 'applied' : 'noop' };
    };
    const host = createHost({ ...m.options, ...(mode === 'opt-out' ? {} : {
      selectionAck: { echoContext: () => mode === 'no-context' ? undefined : echoContext, deferFrame: ack => { deferredAcks.push(ack); } },
    }) });
    const ack = unwrap(await host.handler(null, envelope(1, { kind: mode === 'non-selection' ? 'setReadOnly' : 'selectPoint' })));
    if (mode === 'selection' || mode === 'cross-paragraph' || mode === 'no-context') {
      expect(ack.frame).toBeUndefined();
      expect(ack.version).toBe(0);
      if (mode === 'cross-paragraph') {
        expect(ack.selection?.focus.paraId).toBe('p2');
        expect(ack.editState).toEqual(caretEditState('p2', true));
      }
      expect(ack.echoContext).toEqual(mode === 'no-context' ? undefined : echoContext);
      expect(deferredAcks).toEqual([ack]);
      expect(unwrap(await host.frame()).coversThroughSeq).toBe(1);
    } else {
      expect(ack.frame?.coversThroughSeq).toBe(1);
      expect(ack.echoContext).toBeUndefined();
      expect(deferredAcks).toHaveLength(0);
    }
  }
});

class FakeCommandClient implements DocxCommandClient {
  commandGeneration = 0;
  commandFailure: Error | undefined;
  private readonly initialGenerations = new WeakMap<DocxCommandInitReply, number>();
  readonly calls: DocxCommandEnvelope[] = [];
  readonly replies: DocxCommandAck[] = [];
  readonly initRequests: DocxCommandInitRequest[] = [];
  readonly full = deferred<DocxCommandAck>();
  failureListener: ((error: Error) => void) | null = null;
  readonly frameListeners = new Set<(frame: DocxCommandFrame) => void>();
  fullRequests = 0;
  failureRegistrations = 0;
  disposals = 0;
  editReply?: (command: DocxCommandEnvelope) => Promise<DocxCommandAck>;

  constructor(readonly host?: DocxCommandHost, readonly initial = initialAck()) {}

  async commandInit(request: DocxCommandInitRequest): Promise<DocxCommandInitReply> {
    this.initRequests.push(request);
    const generation = this.commandGeneration;
    const initial = this.host ? unwrap(await this.host.initial()) : this.initial;
    this.initialGenerations.set(initial, generation);
    return initial;
  }

  commandInitialGeneration(initial: DocxCommandInitReply): number | undefined {
    return this.initialGenerations.get(initial);
  }

  async edit(command: DocxCommandEnvelope): Promise<DocxCommandAck> {
    this.calls.push(command);
    if (this.editReply) return this.editReply(command);
    if (!this.host) throw new Error('unexpected edit');
    this.host.policy(command);
    try {
      const reply = unwrap(await this.host.handler(undefined, command));
      this.replies.push(reply);
      return reply;
    } catch (error) {
      if (!isCommandOutcomeUnknownError(error)) throw error;
      throw Object.assign(new Error(error.message), {
        name: error.name, refusal: { ...error.refusal },
      });
    }
  }

  async checkpoint(): Promise<Checkpoint> {
    if (!this.host) throw new Error('unexpected checkpoint');
    this.host.checkpointPolicy();
    return unwrap(await this.host.checkpoint(undefined));
  }

  requestFullFrame(): Promise<DocxCommandAck> {
    this.fullRequests += 1;
    return this.full.promise;
  }

  onFailure(listener: ((error: Error) => void) | null): void {
    this.failureListener = listener;
    if (listener) this.failureRegistrations += 1;
  }

  onFrame(listener: (frame: DocxCommandFrame) => void): () => void {
    this.frameListeners.add(listener);
    return () => { this.frameListeners.delete(listener); };
  }

  destroy(): void { this.disposals += 1; }
}

async function connect(client: DocxCommandClient): Promise<DocxCommandConnection> {
  const pending = createDocxCommandConnection(client, INIT);
  await expect(pending).resolves.toMatchObject({ initial: { session: INIT.session, incarnation: INIT.incarnation } });
  return pending;
}

function sessionOptions(connection: DocxCommandConnection) {
  const queued: (() => void)[] = [];
  const installed: DocxCommandFrame[] = [];
  const retired: string[][] = [];
  const remaining: string[][] = [];
  const dropped: string[] = [];
  const echo: EchoHooks<string, DocxCommandFramePayload> = {
    paint() {},
    drop(entries) { dropped.push(...entries.map(({ value }) => value)); },
    present(frame, done, pending) {
      installed.push(frame);
      retired.push(done.map(({ value }) => value));
      remaining.push(pending.map(({ value }) => value));
    },
  };
  const options: DocxWorkerEditSessionOptions<Uint8Array, string> = {
    source: new Uint8Array([1]), connection, echo, maxInFlight: 2,
    recovery: { reopen: () => Promise.reject(new Error('unexpected recovery')) },
    enqueue: (callback) => { queued.push(callback); },
  };
  return {
    options, installed, retired, remaining, dropped,
    pump() {
      let remaining = 1024;
      while (queued.length) {
        if (--remaining < 0) throw new Error('Command queue did not drain');
        queued.shift()!();
      }
    },
  };
}

test('cross-paragraph early acks publish toolbar state before frames and ignore older frame acks', async () => {
  const m = machine();
  m.state.editState = caretEditState('p1');
  m.hooks.preflight = command => {
    if (command.kind !== 'select' || !('selection' in command.payload) || !command.payload.selection) throw new Error('Unexpected command');
    return { status: 'ready', prepared: {
      command: { ...command, kind: 'select', payload: { selection: command.payload.selection } }, selection: m.state.selection, range: null,
    } };
  };
  m.hooks.apply = ({ command }) => {
    if (command.kind !== 'select' || !command.payload.selection) throw new Error('Missing selection');
    m.state.selection = command.payload.selection;
    m.state.editState = caretEditState(m.state.selection.focus.paraId, m.state.selection.focus.paraId === 'p2');
    return { status: 'noop' };
  };
  const deferredAcks: DocxCommandAck[] = [];
  const host = createHost({ ...m.options,
    selectionAck: { echoContext: () => undefined, deferFrame: ack => { deferredAcks.push(ack); } },
  });
  const client = new FakeCommandClient(host);
  const connection = await connect(client);
  let receiveAck!: (ack: DocxCommandAck) => void;
  connection.client.onAck = listener => { receiveAck = listener; return () => {}; };
  const h = sessionOptions(connection);
  const session = createSession(h.options);
  const published: DocxEditState[] = [];
  const stop = session.subscribe(state => { if (state.editState) published.push(state.editState); });
  const select = (paraId: string) => {
    const at = { story: 'body', paraId, offset: 1 };
    return session.submit('select', { selection: { anchor: at, focus: at } }, { fence: true });
  };
  try {
    const first = select('p2');
    h.pump();
    expect(await first.ack).toMatchObject({ status: 'noop' });
    expect(session.state.editState).toEqual(caretEditState('p2', true));
    expect(published.at(-1)).toEqual(caretEditState('p2', true));
    expect(session.state).toMatchObject({ version: 0, throughSeq: 1, presentedThroughSeq: 0 });
    expect(h.installed).toHaveLength(1);
    expect(deferredAcks).toHaveLength(1);
    const oldFrame = unwrap(await host.frame());
    const second = select('p3');
    h.pump();
    expect(await second.ack).toMatchObject({ status: 'noop' });
    expect(session.state.editState).toEqual(caretEditState('p3'));
    expect(h.installed).toHaveLength(1);
    expect(deferredAcks.map(ack => ack.throughSeq)).toEqual([1, 2]);
    receiveAck({ ...deferredAcks[0]!, frame: oldFrame });
    expect(session.state.selection?.focus.paraId).toBe('p3');
    expect(session.state.editState).toEqual(caretEditState('p3'));
    const currentFrame = unwrap(await host.frame());
    for (const listener of client.frameListeners) listener(currentFrame);
    receiveAck(deferredAcks[0]!);
    expect(session.state).toMatchObject({ version: 0, throughSeq: 2, presentedThroughSeq: 2 });
    expect(session.state.editState).toEqual(caretEditState('p3'));
    expect(await second.presented).toBe('presented');
  } finally { stop(); await session.dispose(); }
});

test('docx sessions publish ack echo contexts after updating selection without presenting a frame', async () => {
  const client = new FakeCommandClient();
  const at = { story: 'body', paraId: 'new', offset: 1 };
  const selection = { anchor: at, focus: at };
  const echoContext: NonNullable<DocxCommandAck['echoContext']> = {
    page: { pageId: '1', pageIndex: 0 }, frameEpoch: 0, layoutRevision: 1, serializedBytes: 100,
    text: 'ab', styleRuns: [], boundaries: [], damage: [], selection: { anchorGrapheme: 1, focusGrapheme: 1 },
    transform: [1, 0, 0, 1, 0, 0], shaping: 'simple-ltr',
  };
  client.editReply = async command => emptyAck({ throughSeq: command.seq, selection, echoContext,
    outcomes: [{ seq: command.seq, status: 'noop' }] });
  const h = sessionOptions(await connect(client));
  const session = createSession(h.options);
  const received: DocxCommandAck[] = [];
  const unsubscribe = session.subscribeAck!(ack => {
    expect(session.state.selection).toEqual(selection);
    expect(session.state.presentedThroughSeq).toBe(0);
    received.push(ack);
  });
  try {
    const ticket = session.submit('navigate', { unit: 'character', direction: 'forward', extend: false }, { fence: true });
    h.pump();
    await ticket.ack;
    expect(received).toHaveLength(1);
    expect(received[0]?.echoContext).toEqual(echoContext);
    unsubscribe();
    const next = session.submit('navigate', { unit: 'character', direction: 'forward', extend: false }, { fence: true });
    h.pump();
    await next.ack;
    expect(received).toHaveLength(1);
  } finally { unsubscribe(); await session.dispose(); }
});

test.each(['stale', 'base-mismatch', 'rejected'] as const)('discarded %s command frames close page and background bitmaps once', async (reason) => {
  const client = new FakeCommandClient();
  const connection = await connect(client);
  const h = sessionOptions(connection);
  const session = createSession(h.options);
  const closed: string[] = [];
  const bitmap = { close() { closed.push('page'); } } as ImageBitmap;
  const background = { close() { closed.push('background'); } } as ImageBitmap;
  const revision = reason === 'stale' ? 0 : 1;
  const frame: DocxCommandFrame = {
    ...connection.initial.frame,
    revision,
    ...(reason === 'base-mismatch' ? { baseRevision: 99 } : {}),
    payload: {
      ...connection.initial.frame.payload,
      stagedPresentation: { frameEpoch: revision, pages: [{
        page: { pageId: '1', pageIndex: 0 }, frameEpoch: revision, bitmap,
        backgrounds: [
          { id: 'page', rect: { x: 0, y: 0, width: 1, height: 1 }, bitmap },
          { id: 'background', rect: { x: 0, y: 0, width: 1, height: 1 }, bitmap: background },
        ],
      }] },
    },
  };
  if (reason === 'rejected') h.options.echo.present = () => false;
  try {
    for (const listener of client.frameListeners) listener(frame);
    for (const listener of client.frameListeners) listener(frame);
    expect(closed).toEqual(['page', 'background']);
    expect(h.installed).toHaveLength(1);
  } finally {
    client.full.resolve(emptyAck());
    await session.dispose();
  }
});

test('a command reply arriving after disposal closes its staged bitmap', async () => {
  const client = new FakeCommandClient();
  const reply = deferred<DocxCommandAck>();
  client.editReply = () => reply.promise;
  const connection = await connect(client);
  const h = sessionOptions(connection);
  const session = createSession(h.options);
  let closed = 0;
  const bitmap = { close() { closed += 1; } } as ImageBitmap;
  const ticket = session.submit('insertText', { text: 'late' });
  h.pump();
  await session.dispose();
  reply.resolve(emptyAck({ throughSeq: 1, version: 1, outcomes: [{ seq: 1, status: 'applied' }], frame: {
    ...connection.initial.frame, revision: 1, coversThroughSeq: 1,
    payload: { ...connection.initial.frame.payload, stagedPresentation: { frameEpoch: 1, pages: [{
      page: { pageId: '1', pageIndex: 0 }, frameEpoch: 1, bitmap,
    }] } },
  } }));
  await reply.promise;
  for (let tick = 0; tick < 8; tick += 1) await Promise.resolve();
  expect(closed).toBe(1);
  expect(h.installed).toHaveLength(1);
  await expect(ticket.presented).resolves.toBe('dropped');
});

function syntheticDocx(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  ));
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function openResident() {
  const { preloadEditWasm } = await import('../wasm/edit');
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  const { residentWorkerFactory } = await import('./__fixtures__/residentWorker');
  const { ResidentEngineWorkerClient } = await import('./residentEngineWorkerClient');
  const startWorker = await residentWorkerFactory();
  const worker = startWorker(7101);
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.open(syntheticDocx(), { generation: 'resident-command-test' });
    expect(worker.sessions).toHaveLength(1);
    const engine = worker.sessions[0]!;
    engine.registerFont(new Uint8Array(readFileSync(resolve(
      import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
    ))));
    engine.layoutDocumentWithRegionsRetained(LAYOUT);
    const read = engine.proposalEngine.readParagraphs({ view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    const point = { story: 'body', paraId: read.paragraphs[0]!.paraId, offset: 4 };
    engine.setSelection(point);
    return { client, engine, point, worker };
  } catch (error) {
    client.destroy();
    throw error;
  }
}

function nativeFrame(engine: ResidentEngineSession, bytes: Uint8Array, coversThroughSeq: number) {
  return buildCommandFrame({
    frame: bytes, metadata: decodeFrameDelta(bytes), coversThroughSeq, caret: engine.residentCaretSnapshot(),
    layoutRevision: 1, documentVersion: engine.proposalEngine.version(),
  });
}

function frameText(frame: RetainedFrame): string {
  return frame.displayList.pages.flatMap(({ primitives }) => primitives)
    .map((primitive) => primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : '').join('');
}

test('d0_frame_packing_uses_supplied_metadata_without_decoding_payload', () => {
  let built!: ReturnType<typeof buildCommandFrame>;
  expect(() => {
    built = buildCommandFrame({
      frame: new Uint8Array([1]), metadata: { full: false, frameEpoch: 17, baseFrameEpoch: 16 },
      coversThroughSeq: 4, caret: { frameEpoch: 17, caretRect: null }, layoutRevision: 1,
    });
  }).not.toThrow();
  expect(built.value).toMatchObject({ revision: 17, baseRevision: 16, coversThroughSeq: 4 });
  expect(new Uint8Array(built.value.payload.frame)).toEqual(new Uint8Array([1]));
});

test('d0_shared_state_comparison_tracks_sticky_bytes_and_editstate', async () => {
  const m = machine();
  m.state.selection.anchor.sticky = new Uint8Array([1, 2]);
  m.state.selection.focus.sticky = new Uint8Array([3, 4]);
  const host = createHost(m.options);
  const initial = unwrap(await host.initial());
  expect(initial.selection?.anchor.sticky).toEqual(new Uint8Array([1, 2]));
  initial.selection!.anchor.sticky![1] = 7;
  const send = async (seq: number) => {
    const command = envelope(seq);
    host.policy(command);
    return unwrap(await host.handler(undefined, command));
  };
  const first = await send(1);
  expect(first).not.toHaveProperty('selection');
  expect(first).not.toHaveProperty('editState');
  m.state.selection = structuredClone(m.state.selection);
  m.state.editState = structuredClone(m.state.editState);
  const same = await send(2);
  expect(same).not.toHaveProperty('selection');
  expect(same).not.toHaveProperty('editState');
  m.state.selection.anchor.sticky![1] = 9;
  m.state.editState = {
    ...m.state.editState, canUndo: true,
    storedMarks: { story: 'body', paraId: 'p1', clear: false, delta: { bold: true } },
  };
  const changed = await send(3);
  expect(changed.selection?.anchor.sticky).toEqual(new Uint8Array([1, 9]));
  expect(changed.editState).toEqual(m.state.editState);
  m.state.selection = structuredClone(m.state.selection);
  m.state.editState = structuredClone(m.state.editState);
  const unchanged = await send(4);
  expect(unchanged).not.toHaveProperty('selection');
  expect(unchanged).not.toHaveProperty('editState');
}, TIMEOUT);

test('d0_recovery_requires_resident_full_frame_capability', async () => {
  type SharedConnection = CommandConnection<
    DocxCommandPayload,
    DocxCommandSelection | null, DocxEditState, DocxCommandFramePayload
  >;
  type Reopened = Awaited<ReturnType<DocxWorkerEditSessionOptions<Uint8Array, string>['recovery']['reopen']>>;
  const acceptsPlainConnection: SharedConnection extends Reopened ? true : false = false;
  expect(acceptsPlainConnection).toBe(false);
  const firstClient = new FakeCommandClient();
  const pending = deferred<DocxCommandAck>();
  firstClient.editReply = () => pending.promise;
  const first = await connect(firstClient);
  const reopenedClient = new FakeCommandClient(undefined, initialAck({ incarnation: 4 }));
  reopenedClient.editReply = async ({ seq }) => emptyAck({
    incarnation: 4, version: 1, throughSeq: seq, outcomes: [{ seq, status: 'applied' }],
    frame: {
      coversThroughSeq: seq, revision: 2, baseRevision: 99,
      payload: { frame: new ArrayBuffer(0), caret: { frameEpoch: 2, caretRect: null }, layoutRevision: 1 },
    },
  });
  const h = sessionOptions(first);
  h.options.recovery.reopen = ({ incarnation }) => createDocxCommandConnection(reopenedClient, { ...INIT, incarnation });
  const session = createSession(h.options);
  try {
    const ticket = session.submit('insertText', { text: 'a' });
    h.pump();
    firstClient.failureListener?.(new SessionFailure('crash', 'worker stopped'));
    await expect(ticket.ack).resolves.toMatchObject({ status: 'applied' });
    expect(reopenedClient.fullRequests).toBe(1);
    reopenedClient.full.resolve(emptyAck({
      incarnation: 4, throughSeq: 1, version: 1,
      frame: {
        coversThroughSeq: 1, revision: 3,
        payload: { frame: new ArrayBuffer(0), caret: { frameEpoch: 3, caretRect: null }, layoutRevision: 1 },
      },
    }));
    await expect(ticket.presented).resolves.toBe('presented');
    expect(session.state).toMatchObject({ version: 1, throughSeq: 1, presentedThroughSeq: 1 });
  } finally {
    pending.resolve(emptyAck());
    await session.dispose();
  }
}, TIMEOUT);

test('d0_epoch_changed_ack_terminates_instead_of_stranding_tickets', async () => {
  const client = new FakeCommandClient();
  client.editReply = async ({ seq }) => emptyAck({
    epoch: INIT.epoch + 1, outcomes: [{ seq, status: 'refused', reason: 'stale-epoch', committed: false }],
  });
  const connection = await connect(client);
  const h = sessionOptions(connection);
  const session = createSession(h.options);
  try {
    const ticket = session.submit('insertText', { text: 'old' });
    const flushed = session.flush();
    h.pump();
    await expect(ticket.ack).rejects.toMatchObject({ code: 'disposed' });
    await expect(flushed).rejects.toMatchObject({ code: 'disposed' });
    await expect(ticket.presented).resolves.toBe('dropped');
    expect(connection.client.failure).toMatchObject({ code: 'disposed' });
    expect(client.calls).toHaveLength(1);
  } finally {
    await session.dispose();
  }
}, TIMEOUT);

test('command_same_coverage_frames_reject_stale_revisions_without_settling_tickets', async () => {
  const client = new FakeCommandClient();
  const connection = await connect(client);
  const h = sessionOptions(connection);
  const session = createSession(h.options);
  try {
    const ticket = session.submit('insertText', { text: 'pending' }, { echo: 'pending' });
    let acknowledged = false;
    let presented = false;
    void ticket.ack.then(() => { acknowledged = true; });
    void ticket.presented.then(() => { presented = true; });
    const push = (frame: DocxCommandFrame) => {
      for (const listener of client.frameListeners) listener(frame);
    };
    const next: DocxCommandFrame = {
      ...connection.initial.frame, revision: 1, baseRevision: 0,
    };
    expect(client.frameListeners.size).toBe(1);
    push(next);
    expect(h.installed.map(({ revision }) => revision)).toEqual([0, 1]);
    expect(h.retired.at(-1)).toEqual([]);
    expect(h.remaining.at(-1)).toEqual(['pending']);
    push({ ...next, baseRevision: 99 });
    push({ ...next, revision: 0, baseRevision: 99 });
    expect(h.installed.map(({ revision }) => revision)).toEqual([0, 1]);
    expect(client.fullRequests).toBe(0);
    expect(client.calls).toEqual([]);
    expect(session.state).toMatchObject({ version: 0, throughSeq: 0, presentedThroughSeq: 0 });
    await Promise.resolve();
    expect(acknowledged).toBe(false);
    expect(presented).toBe(false);
  } finally {
    await session.dispose();
  }
  expect(client.frameListeners.size).toBe(0);
}, TIMEOUT);

test('command_frame_waits_for_admitted_edits_without_blocking_the_docx_host', async () => {
  const m = machine();
  const host = createHost(m.options);
  await host.initial();
  const command = envelope(1);
  host.policy(command);
  let built = false;
  const pushed = host.frame().then((frame) => { built = true; return unwrap(frame); });
  await Promise.resolve();
  expect(built).toBe(false);
  const edit = host.handler(undefined, command);
  expect(unwrap(await edit)).toMatchObject({ throughSeq: 1, version: 1 });
  expect(await pushed).toMatchObject({ coversThroughSeq: 1 });
  expect(m.state.applied).toEqual([1]);
}, TIMEOUT);

test('command_pushed_frames_do_not_overlap_full_frame_recovery', async () => {
  const m = machine();
  const host = createHost(m.options);
  await host.initial();
  const held = deferred<void>();
  const entered = deferred<void>();
  const calls: boolean[] = [];
  const render = m.hooks.frame;
  m.hooks.frame = async (input) => {
    calls.push(input.full);
    if (input.full) { entered.resolve(); await held.promise; }
    return render(input);
  };
  const full = host.requestFullFrame();
  await entered.promise;
  const pushed = host.frame();
  try {
    await Promise.resolve();
    await Promise.resolve();
    expect(calls).toEqual([true]);
    held.resolve();
    const recovered = unwrap(await full);
    const next = unwrap(await pushed);
    expect(calls).toEqual([true, false]);
    expect(next.coversThroughSeq).toBe(recovered.throughSeq);
    expect(next.revision).toBeGreaterThan(recovered.frame!.revision);
  } finally {
    held.resolve();
    await Promise.allSettled([full, pushed]);
  }
}, TIMEOUT);

test('d0_registers_edit_and_checkpoint_policies_once', async () => {
  const m = machine();
  const writes: string[] = [];
  const registry: DocxCommandRegistry = {
    handlers: new Proxy(m.registry.handlers, {
      set(target, key, value: unknown) { writes.push(`handler:${String(key)}`); return Reflect.set(target, key, value); },
    }),
    policies: new Proxy(m.registry.policies, {
      set(target, key, value: unknown) { writes.push(`policy:${String(key)}`); return Reflect.set(target, key, value); },
    }),
  };
  const host = createHost({ ...m.options, registry });
  expect(registry.handlers.edit).toBe(host.handler);
  expect(registry.handlers.checkpoint).toBe(host.checkpoint);
  expect(registry.policies.edit).toBe(host.policy);
  expect(registry.policies.checkpoint).toBe(host.checkpointPolicy);
  const initial = unwrap(await host.initial());
  expect(initial).toMatchObject({ version: 0, throughSeq: 0, outcomes: [], selection: m.state.selection, editState: m.state.editState });
  expect(initial.frame?.coversThroughSeq).toBe(0);
  expect(host.policy(envelope(1))).toMatchObject({ lane: 'input', reorderable: false, mutates: true, reframes: true });
  const first = unwrap(await host.handler(undefined, envelope(1)));
  expect(first).toMatchObject({ throughSeq: 1, version: 1 });
  expect(first).not.toHaveProperty('selection');
  expect(first).not.toHaveProperty('editState');
  expect(first.frame?.coversThroughSeq).toBe(1);
  expect(host.checkpointPolicy()).toMatchObject({ lane: 'input', reorderable: false });
  await host.checkpoint(undefined);
  await host.requestFullFrame();
  expect(writes.sort()).toEqual(['handler:checkpoint', 'handler:edit', 'policy:checkpoint', 'policy:edit']);
  const packed = packCommandAck(first, m.hooks, true);
  expect(unwrap(packed)).toMatchObject({ selection: m.state.selection, editState: m.state.editState });
  const bytes = new ArrayBuffer(4);
  const transferred = packCommandAck(transferable(first, [bytes]), m.hooks, true);
  expect(isTransferResult(transferred)).toBe(true);
  if (!isTransferResult(transferred)) throw new Error('ack wrapper lost transfers');
  expect(transferred.transfer).toEqual([bytes]);
}, TIMEOUT);

test('d0_pipelines_confirmed_basis_and_deduplicates_sequences', async () => {
  const m = machine();
  const gate = deferred<void>();
  const entered = deferred<void>();
  const apply = m.hooks.apply;
  m.hooks.apply = async (prepared, context) => {
    if (prepared.command.seq === 1) { entered.resolve(); await gate.promise; }
    return apply(prepared, context);
  };
  const host = createHost({ ...m.options, retainedOutcomes: 0 });
  const client = new FakeCommandClient(host);
  const connection = await connect(client);
  const h = sessionOptions(connection);
  const session = createSession(h.options);
  try {
    const first = session.submit('insertText', { text: 'a' }, { echo: 'a' });
    const second = session.submit('insertText', { text: 'b' }, { after: first, echo: 'b' });
    h.pump();
    await entered.promise;
    expect(client.calls.map(({ basis }) => basis)).toEqual([{ epoch: 7, version: 0 }, { epoch: 7, version: 0 }]);
    expect(client.calls.map(({ seq, afterSeq }) => [seq, afterSeq])).toEqual([[1, undefined], [2, 1]]);
    gate.resolve();
    await expect(first.ack).resolves.toMatchObject({ seq: 1, status: 'applied' });
    await expect(second.ack).resolves.toMatchObject({ seq: 2, status: 'applied' });
    await expect(second.presented).resolves.toBe('presented');
    expect(m.state.text).toBe('ab');
    expect(m.state.applied).toEqual([1, 2]);
    expect(m.state.contexts[1]).toMatchObject({ currentVersion: 1, basis: { version: 0 }, predecessorSelection: m.state.selection });
    expect(client.replies[1]).not.toHaveProperty('selection');
    expect(client.replies[1]).not.toHaveProperty('editState');
    const duplicate = await connection.client.call.edit(client.calls[0]!);
    expect(duplicate).toMatchObject({ version: 2, throughSeq: 2, outcomes: [{ seq: 1, status: 'duplicate' }], selection: m.state.selection, editState: m.state.editState });
    expect(m.state.applied).toEqual([1, 2]);
    expect(h.retired.at(-1)).toEqual(['a', 'b']);
    expect(h.remaining.at(-1)).toEqual([]);
    expect(h.dropped).toEqual([]);
  } finally {
    gate.resolve();
    await session.dispose();
  }
}, TIMEOUT);

test('d0_refusal_advances_coverage_without_mutation', async () => {
  const m = machine();
  const host = createHost(m.options);
  await host.initial();
  const send = async (command: DocxCommandEnvelope) => {
    host.policy(command);
    return unwrap(await host.handler(undefined, command));
  };
  await send(envelope(1));
  m.state.editState = { ...m.state.editState, readOnly: true };
  const stale = await send(envelope(2, { expectVersion: true }));
  const refused = await send(envelope(3, { payload: { text: 'refuse' }, basis: { epoch: 7, version: 1 } }));
  const cancelled = await send(envelope(4, { afterSeq: 3 }));
  const noop = await send(envelope(5, { payload: { text: '' } }));
  const epoch = await send(envelope(6, { basis: { epoch: 6, version: 1 } }));
  for (const [reply, seq, status, reason] of [
    [stale, 2, 'refused', 'stale-version'], [refused, 3, 'refused', 'locked'],
    [cancelled, 4, 'cancelled', 'dependency-refused'],
  ] as const) {
    expect(reply).toMatchObject({ throughSeq: seq, version: 1, selection: m.state.selection, editState: m.state.editState, outcomes: [{ seq, status, reason, committed: false }] });
  }
  expect(noop).toMatchObject({ throughSeq: 5, version: 1, outcomes: [{ seq: 5, status: 'noop' }] });
  expect(epoch).toMatchObject({
    throughSeq: 5, version: 1, selection: m.state.selection, editState: m.state.editState,
    outcomes: [{ seq: 6, status: 'refused', reason: 'stale-epoch', committed: false }],
  });
  expect(m.state.text).toBe('a');
  expect(m.state.applied).toEqual([1]);
  expect(m.state.preflighted).toEqual([1, 3, 5]);
  const full = unwrap(await host.requestFullFrame());
  expect(full.frame).toMatchObject({ coversThroughSeq: 5 });
  expect(full.frame).not.toHaveProperty('baseRevision');
}, TIMEOUT);

test('d0_binary_delta_base_mismatch_requests_full_frame', async () => {
  createHost(machine().options);
  const arm = await openResident();
  let session: DocxWorkerEditSession<string> | undefined;
  try {
    const firstBytes = arm.engine.buildDisplayListFrame('{}', 0);
    const first = nativeFrame(arm.engine, firstBytes, 0);
    expect(first.value.revision).toBe(decodeFrameDelta(firstBytes).frameEpoch);
    expect(first.value).not.toHaveProperty('baseRevision');
    expect(first.transfer).toContain(first.value.payload.frame);
    const intermediate = arm.engine.applyInput('!', first.value.revision);
    const missed = decodeFrameDelta(intermediate);
    const lastBytes = arm.engine.applyInput('?', missed.frameEpoch);
    const last = nativeFrame(arm.engine, lastBytes, 1);
    expect(decodeFrameDelta(lastBytes).full).toBe(false);
    expect(last.value.baseRevision).toBe(missed.frameEpoch);
    expect(last.value.baseRevision).not.toBe(first.value.revision);
    const padded = new Uint8Array(lastBytes.length + 8);
    padded.set(lastBytes, 4);
    const sliced = buildCommandFrame({
      frame: padded.subarray(4, 4 + lastBytes.length), pageFrames: [intermediate], coversThroughSeq: 1,
      metadata: decodeFrameDelta(lastBytes),
      caret: arm.engine.residentCaretSnapshot(), layoutRevision: 1,
    });
    expect([...new Uint8Array(sliced.value.payload.frame)]).toEqual([...lastBytes]);
    expect(sliced.transfer).toContain(sliced.value.payload.frame);
    expect(sliced.transfer).toContain(sliced.value.payload.pageFrames![0]!);
    expect(new Set(sliced.transfer).size).toBe(sliced.transfer.length);
    const client = new FakeCommandClient(undefined, initialAck({ frame: first.value }));
    client.editReply = async ({ seq }) => emptyAck({ throughSeq: seq, version: 1, outcomes: [{ seq, status: 'applied' }], frame: last.value });
    const connection = await connect(client);
    const h = sessionOptions(connection);
    session = createSession(h.options);
    const ticket = session.submit('insertText', { text: '!?' }, { echo: '!?' });
    h.pump();
    await expect(ticket.ack).resolves.toMatchObject({ status: 'applied' });
    expect(client.fullRequests).toBe(1);
    expect(h.installed.map(({ revision }) => revision)).toEqual([first.value.revision]);
    let presented = false;
    void ticket.presented.then(() => { presented = true; });
    await Promise.resolve();
    expect(presented).toBe(false);
    const fullBytes = arm.engine.buildDisplayListFrame('{}', 0);
    const full = nativeFrame(arm.engine, fullBytes, 1);
    client.full.resolve(emptyAck({ throughSeq: 1, version: 1, frame: full.value }));
    await expect(ticket.presented).resolves.toBe('presented');
    expect(h.retired.at(-1)).toEqual(['!?']);
    const retained = applyFrameDeltaOwned(null, decodeFrameDelta(h.installed.at(-1)!.payload.frame));
    expect(frameText(retained)).toBe('Seed!?');
  } finally {
    await session?.dispose();
    arm.client.destroy();
  }
}, TIMEOUT);

test('d0_checkpoint_fences_edits_and_copies_state', async () => {
  const m = machine();
  const applyGate = deferred<void>();
  const applyEntered = deferred<void>();
  const snapshotGate = deferred<Uint8Array>();
  const snapshotEntered = deferred<void>();
  const order: string[] = [];
  const apply = m.hooks.apply;
  m.hooks.apply = async (prepared, context) => {
    const seq = prepared.command.seq;
    order.push(`edit:${seq}`);
    if (seq === 1) { applyEntered.resolve(); await applyGate.promise; }
    return apply(prepared, context);
  };
  m.hooks.snapshot = () => { order.push('snapshot'); snapshotEntered.resolve(); return snapshotGate.promise; };
  const host = createHost(m.options);
  host.policy(envelope(1));
  const first = host.handler(undefined, envelope(1));
  host.checkpointPolicy();
  const checkpoint = host.checkpoint(undefined);
  host.policy(envelope(2));
  const second = host.handler(undefined, envelope(2));
  try {
    await applyEntered.promise;
    expect(order).toEqual(['edit:1']);
    applyGate.resolve();
    await snapshotEntered.promise;
    expect(m.state.text).toBe('a');
    expect(order).toEqual(['edit:1', 'snapshot']);
    const borrowed = new Uint8Array([99, 11, 22, 33, 99]).subarray(1, 4);
    snapshotGate.resolve(borrowed);
    const reply = await checkpoint;
    const value = unwrap(reply);
    expect(value).toEqual({ state: new Uint8Array([11, 22, 33]), epoch: 7, version: 1, throughSeq: 1 });
    expect(value.state.buffer).not.toBe(borrowed.buffer);
    borrowed.fill(0);
    expect(value.state).toEqual(new Uint8Array([11, 22, 33]));
    expect(isTransferResult(reply)).toBe(true);
    if (!isTransferResult(reply)) throw new Error('checkpoint must transfer its copy');
    expect(reply.transfer).toEqual([value.state.buffer]);
    expect(unwrap(await first)).toMatchObject({ version: 1, throughSeq: 1 });
    expect(unwrap(await second)).toMatchObject({ version: 2, throughSeq: 2 });
    expect(order).toEqual(['edit:1', 'snapshot', 'edit:2']);
    expect(m.state.text).toBe('aa');
  } finally {
    applyGate.resolve();
    snapshotGate.resolve(m.state.snapshot);
    await Promise.allSettled([first, checkpoint, second]);
  }
}, TIMEOUT);

test('d0_connection_propagates_unknown_outcome_and_worker_failure', async () => {
  const client = new FakeCommandClient();
  const connection = await connect(client);
  const m = machine();
  m.hooks.apply = () => { m.state.text += 'committed'; throw new Error('after mutation'); };
  const host = createHost(m.options);
  client.editReply = async (command) => { host.policy(command); return unwrap(await host.handler(undefined, command)); };
  const failures: SessionFailure[] = [];
  const other: SessionFailure[] = [];
  const off = connection.client.onFailure((failure) => { failures.push(failure); });
  connection.client.onFailure((failure) => { other.push(failure); });
  try {
    const error = await connection.client.call.edit(envelope(1)).catch((cause: unknown) => cause);
    expect(isCommandOutcomeUnknownError(error)).toBe(true);
    expect(error).toMatchObject({ name: COMMAND_OUTCOME_UNKNOWN, refusal: { seq: 1 } });
    expect(m.state.text).toBe('committed');
    expect(client.calls).toHaveLength(1);
    expect(connection.client.failure).toBeInstanceOf(SessionFailure);
    expect(failures).toHaveLength(1);
    expect(other).toEqual(failures);
    expect(client.failureRegistrations).toBe(1);
    off();
    await expect(connection.client.call.edit(envelope(2))).rejects.toBeInstanceOf(SessionFailure);
    expect(client.calls).toHaveLength(1);
  } finally {
    await connection.client.dispose();
    await connection.client.dispose();
  }
  expect(client.disposals).toBe(1);
  const crashedClient = new FakeCommandClient();
  const crashed = await connect(crashedClient);
  const pending = deferred<DocxCommandAck>();
  crashedClient.editReply = () => pending.promise;
  const errors: SessionFailure[] = [];
  const remove = crashed.client.onFailure((failure) => { errors.push(failure); });
  const command = crashed.client.call.edit(envelope(1));
  const failure = new SessionFailure('crash', 'worker crashed');
  crashedClient.failureListener?.(failure);
  pending.reject(failure);
  await expect(command).rejects.toBeInstanceOf(SessionFailure);
  expect(crashed.client.failure).toBe(failure);
  expect(errors).toEqual([failure]);
  remove();
  await crashed.client.dispose();
  expect(crashedClient.failureListener).toBeNull();
}, TIMEOUT);

test('d0_initial_frame_and_session_require_no_main_peer', async () => {
  const m = machine();
  let engine!: ResidentEngineSession;
  let base = 0;
  m.hooks.frame = ({ coversThroughSeq, full }) => {
    const bytes = engine.buildDisplayListFrame('{}', full ? 0 : base);
    const built = nativeFrame(engine, bytes, coversThroughSeq);
    base = built.value.revision;
    return built;
  };
  m.hooks.snapshot = () => engine.encodeState();
  m.hooks.apply = ({ command }) => {
    if (command.kind !== 'insertText') throw new Error('unexpected test command');
    engine.applyInput(command.payload.text, base);
    const selection = engine.selection();
    if (!selection) throw new Error('worker lost selection');
    m.state.selection = { anchor: selection.anchor, focus: selection.head };
    return { status: 'applied' };
  };
  const host = createHost(m.options);
  const arm = await openResident();
  engine = arm.engine;
  let session: DocxWorkerEditSession<string> | undefined;
  try {
    m.state.selection = { anchor: arm.point, focus: arm.point };
    const connection = await connect(new FakeCommandClient(host));
    expect(connection.initial.selection).toEqual(m.state.selection);
    expect(connection.initial.editState).toEqual(m.state.editState);
    const firstFrame = connection.initial.frame!;
    expect(firstFrame.coversThroughSeq).toBe(0);
    expect(firstFrame.payload.frame.byteLength).toBeGreaterThan(0);
    const first = applyFrameDeltaOwned(null, decodeFrameDelta(firstFrame.payload.frame));
    expect(frameText(first)).toBe('Seed');
    const h = sessionOptions(connection);
    session = createSession(h.options);
    expect(session.backend).toBe('worker');
    expect(session.state.presentedThroughSeq).toBe(0);
    expect(h.installed).toHaveLength(1);
    const ticket = session.submit('insertText', { text: '!' });
    h.pump();
    await expect(ticket.ack).resolves.toMatchObject({ seq: 1, status: 'applied' });
    await expect(ticket.presented).resolves.toBe('presented');
    expect(session.state).toMatchObject({ version: 1, throughSeq: 1, presentedThroughSeq: 1 });
    const final = applyFrameDeltaOwned(first, decodeFrameDelta(h.installed.at(-1)!.payload.frame));
    expect(frameText(final)).toBe('Seed!');
    expect(arm.worker.sessions).toHaveLength(1);
    const source = readFileSync(resolve(import.meta.dir, 'createDocxWorkerEditSession.ts'), 'utf8');
    expect(source).not.toMatch(/(?:createYrsSession|bootstrapWorkerOpenPeer|workerOpenReplica|preloadEditWasm|editWasmModule|createEditSession)/);
    expect(source).not.toMatch(/from\s+['"][^'"]*(?:\/wasm(?:\/|['"])|\.\/index['"])/);
    const forbidden: string[] = [];
    const bundle = await Bun.build({
      entrypoints: [resolve(import.meta.dir, 'createDocxWorkerEditSession.ts')],
      target: 'bun',
      plugins: [{
        name: 'docx-command-peer-guard',
        setup(build) {
          build.onResolve({ filter: /.*/ }, ({ path, importer }) => {
            const dependency = path.startsWith('.') ? resolve(dirname(importer), path) : path;
            if (/\/yrs\/(?:index(?:\.ts)?$|wasm\/|residentEngineSession|bootstrapWorkerOpenPeer|workerOpenReplica)/.test(dependency) ||
              /\/docx\/src\/wasm\/(?:edit|generated\/edit)/.test(dependency)) forbidden.push(dependency);
            return undefined;
          });
        },
      }],
    });
    expect(bundle.success).toBe(true);
    expect(forbidden).toEqual([]);
  } finally {
    await session?.dispose();
    arm.client.destroy();
  }
}, TIMEOUT);

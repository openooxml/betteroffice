import {
  createSessionClient, requestWasmCompile, type SessionClient,
} from '../../../../shared/office-session/client';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure, type MethodPolicy, type Promisified } from '../../../../shared/office-session/types';
import { wasmAssetUrl } from '../wasm/asset';
import {
  createWorkbookSnapshotBuilder, initWasm, openWorkbookPeer,
  type OpenWorkbookOptions, type Viewport, type WorkbookCalculationContext, type WorkbookHandle,
} from '../wasm/loader';
import {
  WORKBOOK_SESSION_METHODS,
  WORKBOOK_SESSION_POLICIES,
  WorkbookPngUnavailableError,
  type WorkbookFrame,
  type WorkbookFrameOptions,
  type WorkbookInitialPaint,
  type WorkbookSessionEvents,
  type WorkbookSessionMethods,
  type WorkbookSessionOpenOptions,
  type WorkbookSessionOpenReply,
  type WorkbookSessionState,
  type WorkbookWireFrame,
} from './methods';
import {
  WORKBOOK_INTERNAL_SESSION_METHODS,
  workbookSessionInternals,
  type WorkbookInternalSessionMethods,
  type WorkbookReplayEnvelope,
} from './replay';
import { WorkbookPeerHydrationError } from './peerHydrationError';
import { workbookPeerSources as peerSources, type WorkbookPeerSource } from './clientInternals';
import type {
  CommandBasis, CommandIdentity, CommandStatus, WorkbookAck, WorkbookCommand,
  WorkbookCommandContext, WorkbookCommandFrame, WorkbookCommandSelection, WorkbookEditState,
  WorkbookSharedAck, WorkbookSharedWireAck,
} from './commandProtocol';
import { createWorkbookEditSession, type WorkbookEditSession } from './editSession';
import { createAsyncCollaborationReplica, type AsyncReplicaAdapter } from '../collaboration/asyncReplica';
import { localNowSerial } from './calculationClock';
import { createWorkbookSessionRecovery, workbookSessionRecovery } from './clientRecovery';

type Events = { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] } & {
  peerOpened: { version: string; initialCalculation?: WorkbookCalculationContext | null };
};
type Methods = WorkbookSessionMethods & WorkbookInternalSessionMethods;
const wasmModules = new Map<string, WebAssembly.Module>();
const PEER_SNAPSHOT_BUDGET = { records: 256, bytes: 16 * 1024 } as const;
const PEER_SNAPSHOT_SLICE_MS = 8;

/**
 * Options for opening a workbook in a dedicated worker.
 * @experimental
 */
export interface OpenWorkbookSessionOptions extends OpenWorkbookOptions {
  worker?: () => Worker;
  wasm?: ArrayBuffer | WebAssembly.Module;
  /** Retain source bytes and a compiled module for edit peer hydration. */
  retainPeerHydration?: boolean;
  /** Aborting closes the session worker while the open is still in flight. */
  signal?: AbortSignal;
  initialViewport?: WorkbookSessionOpenOptions['initialViewport'];
  includeEditState?: boolean;
  initialUpdate?: Uint8Array;
}

/**
 * Async workbook access and its current local projection.
 * @experimental
 */
export interface WorkbookSession {
  readonly pngExport?: boolean;
  readonly replica?: AsyncReplicaAdapter;
  bindCollaboration?(apply: (update: Uint8Array) => Promise<unknown>): () => void;
  readonly state: WorkbookSessionState;
  readonly initialPaint?: WorkbookInitialPaint;
  readonly commandIdentity: CommandIdentity;
  readonly confirmedBasis: CommandBasis;
  readonly throughSeq: number;
  readonly presentedThroughSeq: number;
  readonly editState: Readonly<Partial<WorkbookEditState>>;
  readonly selection: WorkbookCommandSelection | undefined;
  readonly editSession?: WorkbookEditSession;
  submit(command: WorkbookCommand): Promise<CommandStatus>;
  onAck(listener: (ack: WorkbookAck) => void): () => void;
  markPresented(frame: WorkbookCommandFrame | Pick<WorkbookCommandFrame, 'session' | 'incarnation' | 'coversThroughSeq'>): boolean;
  whenPresented(seq: number): Promise<void>;
  readonly call: Promisified<Omit<WorkbookSessionMethods,
    'open' | 'openRecovery' | 'dispose' | 'frame' | 'command' | 'edit' | 'checkpoint' | 'commandFrame' |
    'renderPng' | 'renderRangePng' | 'printDisplayList'>> & {
    renderPng(viewport: Viewport, options?: WorkbookFrameOptions): Promise<Uint8Array>;
    renderRangePng(options?: { range?: string; scale?: number }, frameOptions?: WorkbookFrameOptions): Promise<Uint8Array>;
    printDisplayList(...args: Parameters<WorkbookHandle['printDisplayList']>): Promise<ReturnType<WorkbookHandle['printDisplayList']>>;
    command(command: WorkbookCommand): Promise<WorkbookAck>;
    /**
     * @experimental A newer `frame` call replaces a queued one, which rejects with an error
     * named `SessionSuperseded`.
     */
    frame(viewport: Viewport, options?: WorkbookFrameOptions): Promise<WorkbookFrame>;
  };
  save(): Promise<Uint8Array>;
  on<K extends keyof WorkbookSessionEvents>(
    name: K, listener: (payload: WorkbookSessionEvents[K]) => void
  ): () => void;
  onFailure(listener: (failure: SessionFailure) => void): () => void;
  readonly failure: SessionFailure | undefined;
  dispose(): Promise<void>;
}

function copyBytes(bytes: Uint8Array | ArrayBuffer): Uint8Array<ArrayBuffer> {
  return ArrayBuffer.isView(bytes)
    ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength).slice()
    : new Uint8Array(bytes).slice();
}

function decodeFrame(frame: WorkbookWireFrame): WorkbookFrame {
  return {
    ...frame,
    displayList: JSON.parse(new TextDecoder().decode(frame.displayList)) as WorkbookFrame['displayList'],
  };
}

async function pngBytes(render: Promise<ArrayBuffer>): Promise<Uint8Array> {
  try { return new Uint8Array(await render); }
  catch (error) {
    if (error instanceof Error && error.name === 'WorkbookPngUnavailableError') throw new WorkbookPngUnavailableError();
    throw error;
  }
}

function prepareOpen(bytes: Uint8Array | ArrayBuffer, options: OpenWorkbookSessionOptions): {
  document: ArrayBuffer;
  input: WorkbookSessionOpenOptions & { retainPeerHydration?: boolean };
  transfer: Transferable[];
} {
  const document = copyBytes(bytes).buffer;
  const transfer: Transferable[] = [document];
  const calculation = options.calculation ?? (options.includeEditState && !options.retainPeerHydration && !options.collaborative
    ? { nowSerial: localNowSerial(Date.now()), randSeed: globalThis.crypto.getRandomValues(new Uint32Array(1))[0] >>> 0 }
    : undefined);
  const input: WorkbookSessionOpenOptions & { retainPeerHydration?: boolean } = {
    collaborative: options.collaborative,
    clientId: options.clientId,
    calculation: calculation ? { ...calculation } : undefined,
    retainPeerHydration: options.retainPeerHydration,
    initialViewport: options.initialViewport ? { ...options.initialViewport } : undefined,
    includeEditState: options.includeEditState,
    initialUpdate: options.initialUpdate?.slice(),
  };
  if (options.wasm instanceof ArrayBuffer ||
    Object.prototype.toString.call(options.wasm) === '[object ArrayBuffer]') {
    input.wasm = copyBytes(options.wasm as ArrayBuffer).buffer;
    transfer.push(input.wasm);
  } else if (options.wasm !== undefined) input.wasm = options.wasm;
  else if (!options.worker) input.wasm = wasmModules.get(wasmAssetUrl().href);
  return { document, input, transfer };
}

async function openPeerFromSource(source: WorkbookPeerSource): Promise<WorkbookHandle> {
  if (!source.module) throw new WorkbookPeerHydrationError('missing-module',
    'Workbook peer hydration requires a worker-compiled module');
  await initWasm(source.module);
  if (source.disposed) throw new SessionFailure('disposed', 'Workbook session was disposed');
  if (source.hydration === undefined) throw new WorkbookPeerHydrationError('missing-hydration',
    'Workbook worker did not retain peer calculation state');
  if (source.snapshot) {
    const snapshot = source.snapshot;
    let builder: ReturnType<typeof createWorkbookSnapshotBuilder> | undefined;
    let peer: WorkbookHandle | undefined;
    try {
      const pinned = await snapshot.beginPeerSnapshot(PEER_SNAPSHOT_BUDGET.records, PEER_SNAPSHOT_BUDGET.bytes);
      if (pinned.version !== source.version || pinned.sequence !== source.sequence) {
        throw new Error('Workbook snapshot differs from retained peer hydration');
      }
      builder = createWorkbookSnapshotBuilder(source.options);
      for (;;) {
        if (source.disposed) throw new SessionFailure('disposed', 'Workbook session was disposed');
        const chunks = await snapshot.pullPeerSnapshot();
        if (chunks === undefined) break;
        for (const chunk of chunks) builder.push(new Uint8Array(chunk));
      }
      let ready = false;
      for (;;) {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
        const deadline = performance.now() + PEER_SNAPSHOT_SLICE_MS;
        do {
          if (source.disposed) throw new SessionFailure('disposed', 'Workbook session was disposed');
          ready = builder.advance(PEER_SNAPSHOT_BUDGET.records, PEER_SNAPSHOT_BUDGET.bytes);
        } while (!ready && performance.now() < deadline);
        if (ready) break;
      }
      peer = builder.finish();
      if (peer.version() !== pinned.version) throw new Error('Workbook snapshot peer version differs from worker');
      await snapshot.endPeerSnapshot(false);
      if (source.disposed) throw new SessionFailure('disposed', 'Workbook session was disposed');
      source.bytes = undefined;
      source.hydration = undefined;
      return peer;
    } catch (error) {
      try { peer?.dispose(); } catch {}
      try { await snapshot.endPeerSnapshot(true); } catch {}
      if (source.disposed) throw new SessionFailure('disposed', 'Workbook session was disposed');
      try {
        console.warn(`xlsx worker editor: snapshot hydration fell back: ${error instanceof Error ? error.message : String(error)}`);
      } catch {}
    } finally {
      try { builder?.dispose(); } catch {}
    }
  }
  if (source.bytes === undefined || source.hydration === undefined) {
    throw new WorkbookPeerHydrationError('missing-hydration',
      'Workbook worker did not retain peer source bytes or calculation state');
  }
  try { return openWorkbookPeer(source.bytes, source.options, source.hydration); } catch (error) {
    throw new WorkbookPeerHydrationError('missing-hydration',
      `Workbook worker supplied invalid peer hydration: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/** @internal */
export function hydratePeer(session: WorkbookSession): Promise<WorkbookHandle> {
  const source = peerSources.get(session);
  if (!source) return Promise.reject(new TypeError('Workbook session does not retain peer hydration state'));
  if (source.disposed) return Promise.reject(new SessionFailure('disposed', 'Workbook session was disposed'));
  if (!source.pending) {
    const pending = openPeerFromSource(source);
    source.pending = pending;
    void pending.catch(() => {
      if (source.pending === pending) source.pending = undefined;
    });
  }
  return source.pending;
}

/**
 * Opens a worker session, transferring copies so caller buffers stay usable.
 * @experimental
 */
export async function openWorkbookSession(
  bytes: Uint8Array | ArrayBuffer,
  options: OpenWorkbookSessionOptions = {}
): Promise<WorkbookSession> {
  const transport = createWorkerTransport(
    options.worker ? options.worker() :
      new Worker(new URL('./xlsxSessionWorker.mjs', import.meta.url), { type: 'module' })
  );
  if (!options.worker && options.wasm === undefined && !wasmModules.has(wasmAssetUrl().href)) {
    requestWasmCompile(transport);
  }
  return createWorkbookSession(bytes, options, transport);
}

export async function createWorkbookSession(
  bytes: Uint8Array | ArrayBuffer,
  options: OpenWorkbookSessionOptions,
  transport: SessionTransport
): Promise<WorkbookSession> {
  let document: ArrayBuffer;
  let input: WorkbookSessionOpenOptions;
  let transfer: Transferable[];
  let client: SessionClient<Methods, Events>;
  let peerSource: WorkbookPeerSource | undefined;
  let rejectHydration: ((error: WorkbookPeerHydrationError) => void) | undefined;
  let hydrationFailure: Promise<never> | undefined;
  let recoverySource: Uint8Array | undefined;
  try {
    ({ document, input, transfer } = prepareOpen(bytes, options));
    if (options.includeEditState && !options.retainPeerHydration && !input.collaborative) {
      recoverySource = new Uint8Array(document).slice();
    }
    if (options.retainPeerHydration) {
      hydrationFailure = new Promise<never>((_, reject) => { rejectHydration = reject; });
      void hydrationFailure.catch(() => {});
      peerSource = {
        bytes: new Uint8Array(document).slice(),
        options: {
          collaborative: input.collaborative, clientId: input.clientId, calculation: input.calculation,
        },
        module: input.wasm instanceof WebAssembly.Module ? input.wasm : undefined,
        disposed: false,
      };
    }
    client = createSessionClient<Methods, Events>(transport, {
      methods: { ...WORKBOOK_SESSION_METHODS, ...WORKBOOK_INTERNAL_SESSION_METHODS },
      onWasmModuleMessage: ({ url, module, hydration, version, sequence }) => {
        if (url !== wasmAssetUrl().href) return;
        if (peerSource) {
          if (hydration === undefined) {
            rejectHydration!(new WorkbookPeerHydrationError('missing-hydration',
              'Workbook worker did not retain peer calculation state'));
            return;
          }
          if (typeof version !== 'string' || typeof sequence !== 'number' ||
            !Number.isSafeInteger(sequence) || sequence < 0) {
            rejectHydration!(new WorkbookPeerHydrationError('missing-hydration',
              'Workbook worker hydration is missing its committed version or sequence'));
            return;
          }
          peerSource.module = module;
          peerSource.hydration = hydration;
          peerSource.version = version;
          peerSource.sequence = sequence;
        }
        if (options.wasm === undefined && !options.worker && !wasmModules.has(url)) {
          wasmModules.set(url, module);
        }
      },
    });
  } catch (error) {
    try { transport.close(); } catch {}
    throw error;
  }
  let state: WorkbookSessionState;
  let commandContext: WorkbookCommandContext | undefined;
  let initialVersion: string | undefined;
  let initialPaint: WorkbookInitialPaint | undefined;
  let openedEditState: WorkbookEditState | undefined;
  let openedSelection: WorkbookCommandSelection | undefined;
  let pngExport = false;
  let clientId: number | undefined;
  if (peerSource) client.on('peerOpened', (opened) => {
    initialVersion = opened.version;
    if (peerSource) peerSource.initialCalculation = opened.initialCalculation;
  });
  client.on('changed', (change) => { state = { ...state, ...change }; });
  client.onFailure(() => { state = { ...state, stage: 'failed' }; });
  const signal = options.signal;
  const abort = () => {
    if (peerSource) peerSource.disposed = true;
    void client.dispose().catch(() => {});
    try { transport.close(); } catch {}
  };
  try {
    signal?.addEventListener('abort', abort, { once: true });
    if (signal?.aborted) abort();
    const opening = client.callWithTransfer('open', [document, input], transfer);
    const reply: WorkbookSessionOpenReply = await (hydrationFailure ? Promise.race([opening, hydrationFailure]) : opening);
    const { commandContext: openedCommands, initialPaint: paint, clientId: replicaId,
      editState: edits, selection: selected, pngExport: png, ...projection } = reply;
    pngExport = png === true;
    if (paint) initialPaint = { ...paint, frame: { ...paint.frame, ...decodeFrame(paint.frame) } };
    openedEditState = edits ?? paint?.editState;
    openedSelection = selected;
    clientId = replicaId;
    commandContext = openedCommands;
    initialVersion ??= openedCommands?.basis.version;
    state = projection;
    if (signal?.aborted) throw new SessionFailure('disposed', 'Session was disposed');
    if (peerSource && !peerSource.module) {
      throw new WorkbookPeerHydrationError('missing-module',
        'Workbook worker did not retain a compiled module for peer hydration');
    }
    if (peerSource && peerSource.hydration === undefined) {
      throw new WorkbookPeerHydrationError('missing-hydration',
        'Workbook worker did not retain peer calculation state');
    }
  } catch (error) {
    await client.dispose();
    throw error;
  } finally { signal?.removeEventListener('abort', abort); }

  const calls = { ...client.call };
  for (const method of Object.keys(WORKBOOK_SESSION_METHODS) as (keyof WorkbookSessionMethods)[]) {
    Object.defineProperty(calls, method, { value: (...args: unknown[]) => {
      if (method === 'command' && (peerSource || workbookSessionInternals.get(session)?.editPeerAttached)) {
        return Promise.reject(new Error('Workbook commands require a session without an edit peer'));
      }
      if (method === 'applyEdits' && commandMode) {
        return Promise.reject(new Error('Workbook command sessions require command batches'));
      }
      if (peerSource || workbookSessionInternals.get(session)?.editPeerAttached) {
        const configured = WORKBOOK_SESSION_POLICIES[method];
        const policy = typeof configured === 'function'
          ? (configured as (...args: unknown[]) => MethodPolicy)(...args) : configured;
        if (policy.mutates) {
          if (peerSource) {
            return Promise.reject(new WorkbookPeerHydrationError('mutation-outside-replay',
              `Cannot call ${method} outside workbook edit-peer replay in retained hydration mode`));
          }
          return Promise.reject(new Error(`Cannot call ${method} while a workbook edit peer is attached`));
        }
      }
      const call = client.call[method] as (...args: unknown[]) => Promise<unknown>;
      try {
        if (method === 'renderPng' || method === 'renderRangePng') {
          const options = args[1] as WorkbookFrameOptions | undefined;
          args = structuredClone([args[0], { ...options, sheet: options?.sheet ?? state.activeSheet }]);
        } else if (method === 'exportStructured' || method === 'exportMarkdown' || method === 'printDisplayList') {
          args = structuredClone(args);
        }
      } catch (error) { return Promise.reject(error); }
      return edits.runAfterEdits(() => call(...args));
    } });
  }
  const {
    version, readCells, findText, validateEdits, applyEdits, frame, sheetView, cellGeometry,
    cellInputs, captureFormat, sheets, calculationStatus, save,
    renderPng, renderRangePng, exportStructured, exportMarkdown, printDisplayList,
    collaborationStateVector, collaborationStateAsUpdate,
  } = calls;
  const commandIdentity: CommandIdentity = Object.freeze(commandContext ? {
    session: commandContext.session, incarnation: commandContext.incarnation,
  } : { session: '', incarnation: '' });
  let commandMode = false;
  let selectionPublished = openedSelection !== undefined;
  const basis = commandContext?.basis ?? { epoch: 1, version: initialVersion ?? '' };
  const recovery = recoverySource ? createWorkbookSessionRecovery(recoverySource,
    { ...options, calculation: input.calculation },
    async (allowed) => {
      const dispatched = await edits.runAfterEdits(() => Promise.resolve({
        reply: allowed() ? client.call.checkpoint(true) : undefined,
      }));
      return dispatched.reply;
    }, basis.epoch, options.calculation === undefined, initialVersion) : undefined;
  recoverySource = undefined;
  function decodeAck(reply: WorkbookSharedWireAck): WorkbookSharedAck {
    const { frame: paint, compatibility, ...receipt } = reply;
    const decoded: WorkbookSharedAck = { ...receipt,
      ...(paint ? { frame: { ...paint, payload: { ...paint.payload, ...decodeFrame(paint.payload) } } } : {}) };
    if (compatibility) {
      const { frame, ...status } = compatibility;
      decoded.compatibility = { ...status,
        ...(frame ? { frame: { ...frame, ...decodeFrame(frame) } } : {}) };
    }
    return decoded;
  }
  const edits: WorkbookEditSession = createWorkbookEditSession({ identity: commandIdentity,
    onAdmit: recovery ? (input) => recovery.admit(input) : undefined,
    normalizePayload: recovery ? (payload) => recovery.normalizePayload(payload) : undefined,
    connection: {
      initial: { session: commandIdentity.session, incarnation: 1, epoch: basis.epoch,
        throughSeq: 0, version: 0, engineVersion: basis.version, outcomes: [],
        ...(openedEditState ? { editState: openedEditState } : {}),
        ...(openedSelection ? { selection: openedSelection } : {}),
        ...(initialPaint ? { frame: { coversThroughSeq: 0, revision: initialPaint.frame.epoch,
          payload: initialPaint.frame } } : {}) },
      client: {
        call: { edit: async (input) => {
          const compatibility = input.payload.compatibility;
          if (!compatibility || compatibility.kind === 'applyUpdate') return decodeAck(await client.call.edit(input));
          const reply = await client.call.command(compatibility);
          if (!reply.shared) throw new Error('Workbook command omitted shared acknowledgement');
          return decodeAck(reply.shared);
        },
          checkpoint: () => client.call.checkpoint() },
        get failure() { return client.failure; },
        onFailure: (listener) => client.onFailure(listener),
        dispose: () => client.dispose(),
      },
    },
    enqueue: (callback) => { callback(); },
    requestFullFrame: async () => {
      const paint = await client.call.commandFrame();
      return { session: commandIdentity.session, incarnation: 1, epoch: basis.epoch,
        throughSeq: edits.state.throughSeq, version: edits.state.version, engineVersion: edits.engineVersion,
        outcomes: [], ...(paint ? { frame: { ...paint, payload: { ...paint.payload, ...decodeFrame(paint.payload) } } } : {}) };
    },
  });
  edits.onFailure(() => { state = { ...state, stage: 'failed' }; });
  edits.onAck((ack) => {
    recovery?.acknowledge(ack);
    if (ack.selection && (selectionPublished ||
      ack.statuses.some((status) => status.status === 'applied' || status.status === 'noop'))) selectionPublished = true;
  });
  edits.subscribe((current) => {
    if (current.editState) state = { ...state,
      ...(current.editState.activeSheet === undefined ? {} : { activeSheet: current.editState.activeSheet }),
      ...(current.editState.sheets === undefined ? {} : { sheets: current.editState.sheets }) };
  });
  async function command(input: WorkbookCommand): Promise<WorkbookAck> {
    commandMode = true;
    if (peerSource || workbookSessionInternals.get(session)?.editPeerAttached) {
      throw new Error('Workbook commands require a session without an edit peer');
    }
    return edits.command(input);
  }
  const session: WorkbookSession = {
    pngExport,
    get state() { return state; },
    initialPaint,
    commandIdentity,
    get confirmedBasis(): CommandBasis { return Object.freeze({ epoch: edits.state.epoch, version: edits.engineVersion }); },
    get throughSeq() { return edits.state.throughSeq; },
    get presentedThroughSeq() { return Math.max(0, edits.state.presentedThroughSeq); },
    get editState() { return edits.state.editState ?? {}; },
    get selection() { return selectionPublished ? edits.state.selection : undefined; },
    editSession: edits,
    submit: async (input) => {
      const ack = await command(input);
      const status = ack.statuses.find((entry) => entry.seq === input.seq);
      if (!status) throw new Error(`Workbook acknowledgement omitted sequence ${input.seq}`);
      return status;
    },
    onAck: (listener) => edits.onAck(listener),
    markPresented: (paint) => edits.markPresented(paint),
    whenPresented: async (seq) => {
      if (!Number.isSafeInteger(seq) || seq < 0) return Promise.reject(new RangeError('Invalid presentation sequence'));
      if (client.failure || edits.failure) throw client.failure ?? edits.failure;
      if (seq === 0 && !initialPaint) return;
      if (await edits.whenPresented(seq) !== 'presented') {
        throw client.failure ?? edits.failure ?? new SessionFailure('disposed', 'Presentation was dropped');
      }
    },
    call: {
      version, readCells, findText, validateEdits, applyEdits, sheetView, cellGeometry, cellInputs,
      captureFormat, sheets, calculationStatus, save, command, exportStructured, exportMarkdown,
      collaborationStateVector, collaborationStateAsUpdate,
      renderPng: (viewport, options) => pngBytes(renderPng(viewport, options)),
      renderRangePng: (options, frameOptions) => pngBytes(renderRangePng(options, frameOptions)),
      printDisplayList: async (...args) => JSON.parse(new TextDecoder().decode(await printDisplayList(...args))),
      frame: async (viewport, options) => decodeFrame(await frame(viewport, options)),
    },
    save: async () => new Uint8Array(await save()),
    on: (name, listener) => client.on(name, listener),
    onFailure: (listener) => edits.onFailure(listener),
    get failure() { return edits.failure ?? client.failure; },
    dispose: () => {
      if (peerSource) peerSource.disposed = true;
      recovery?.dispose();
      return edits.dispose();
    },
  };
  if (clientId !== undefined && !peerSource) {
    let applyRemote: ((update: Uint8Array) => Promise<unknown>) | undefined;
    const replica = createAsyncCollaborationReplica({ clientId,
      encodeStateVector: () => collaborationStateVector(),
      encodeStateAsUpdate: (remoteStateVector) => collaborationStateAsUpdate(remoteStateVector?.slice()),
      applyUpdate: async (update) => {
        if (applyRemote) return applyRemote(update.slice());
        const ticket = edits.submit('applyUpdate', { payload: { update: update.slice() } }, { lane: 'collab', fence: true });
        const ack = await ticket.ack;
        if (ack.status !== 'applied' && ack.status !== 'noop') throw new Error(ack.reason ?? 'Remote update refused');
        return ack;
      },
      onUpdate: (listener) => client.on('collaborationUpdate', ({ update, origin }) => listener(update.slice(), origin)),
    });
    Object.defineProperty(session, 'replica', { value: replica, enumerable: true });
    session.bindCollaboration = (apply) => {
      applyRemote = apply;
      return () => { if (applyRemote === apply) applyRemote = undefined; };
    };
  }
  const replay = async (envelope: WorkbookReplayEnvelope) => {
    const reply = await client.call.replay(envelope).catch((error: unknown) => {
      if (error instanceof Error && error.name === 'WorkbookPeerHydrationError') {
        throw new WorkbookPeerHydrationError('mutation-outside-replay', error.message);
      }
      throw error;
    });
    if (envelope.op.method === 'setActiveSheet') {
      state = { ...state, activeSheet: envelope.op.args[0] };
    } else if (reply.result !== null && typeof reply.result === 'object' && 'sheetInfo' in reply.result) {
      const info = reply.result.sheetInfo;
      state = {
        ...state, activeSheet: info.activeSheet,
        sheets: info.sheetIds.map((id, index) => ({ id, index, name: info.sheetNames[index] })),
      };
    }
    return reply;
  };
  const source = peerSource;
  if (source) source.snapshot = {
    beginPeerSnapshot: (...args) => client.call.beginPeerSnapshot(...args),
    pullPeerSnapshot: () => client.call.pullPeerSnapshot(),
    endPeerSnapshot: (discard) => client.call.endPeerSnapshot(discard),
  };
  const attachPeer = source ? async (peerVersion: string) => {
    try {
      await client.call.attachPeer(peerVersion, source.sequence!);
    } catch (error) {
      if (error instanceof Error && error.name === 'WorkbookPeerHydrationError') {
        throw new WorkbookPeerHydrationError('version-mismatch', error.message);
      }
      throw error;
    }
  } : undefined;
  workbookSessionInternals.set(session, {
    replay, initialVersion: peerSource?.version ?? initialVersion, attachPeer,
    detachPeer: source ? () => client.call.detachPeer() : undefined, editPeerAttached: false,
    cellInput: source ? async (sheet, row, col) => {
      let column = '';
      for (let at = col + 1; at > 0; at = Math.floor((at - 1) / 26)) {
        column = String.fromCharCode(65 + (at - 1) % 26) + column;
      }
      return (await client.call.cellInputs(sheet, `${column}${row + 1}`)).cells[0]?.[0]?.input ?? '';
    } : undefined,
    preview: source ? async (viewport, sheet, ops) => decodeFrame(await client.call.preview(viewport, sheet, ops)) : undefined,
    get initialCalculation() { return source?.initialCalculation; },
  });
  if (peerSource) peerSources.set(session, peerSource);
  if (recovery) workbookSessionRecovery.set(session, recovery.recovery);
  return session;
}

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
  type WorkbookFrame,
  type WorkbookFrameOptions,
  type WorkbookSessionEvents,
  type WorkbookSessionMethods,
  type WorkbookSessionOpenOptions,
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
}

/**
 * Async workbook access and its current local projection.
 * @experimental
 */
export interface WorkbookSession {
  readonly state: WorkbookSessionState;
  readonly call: Promisified<Omit<WorkbookSessionMethods, 'open' | 'dispose' | 'frame'>> & {
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

function prepareOpen(bytes: Uint8Array | ArrayBuffer, options: OpenWorkbookSessionOptions): {
  document: ArrayBuffer;
  input: WorkbookSessionOpenOptions & { retainPeerHydration?: boolean };
  transfer: Transferable[];
} {
  const document = copyBytes(bytes).buffer;
  const transfer: Transferable[] = [document];
  const input: WorkbookSessionOpenOptions & { retainPeerHydration?: boolean } = {
    collaborative: options.collaborative,
    clientId: options.clientId,
    calculation: options.calculation === undefined ? undefined : { ...options.calculation },
    retainPeerHydration: options.retainPeerHydration,
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
  let rejectHydration!: (error: WorkbookPeerHydrationError) => void;
  const hydrationFailure = new Promise<never>((_, reject) => { rejectHydration = reject; });
  void hydrationFailure.catch(() => {});
  try {
    ({ document, input, transfer } = prepareOpen(bytes, options));
    if (options.retainPeerHydration) {
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
            rejectHydration(new WorkbookPeerHydrationError('missing-hydration',
              'Workbook worker did not retain peer calculation state'));
            return;
          }
          if (typeof version !== 'string' || typeof sequence !== 'number' ||
            !Number.isSafeInteger(sequence) || sequence < 0) {
            rejectHydration(new WorkbookPeerHydrationError('missing-hydration',
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
  let initialVersion: string | undefined;
  client.on('peerOpened', (opened) => {
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
    state = await Promise.race([
      client.callWithTransfer('open', [document, input], transfer), hydrationFailure,
    ]);
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
      return call(...args);
    } });
  }
  const {
    version, readCells, findText, validateEdits, applyEdits, frame, sheetView, cellGeometry,
    cellInputs, sheets, calculationStatus, save,
  } = calls;
  const session: WorkbookSession = {
    get state() { return state; },
    call: {
      version, readCells, findText, validateEdits, applyEdits, sheetView, cellGeometry, cellInputs,
      sheets, calculationStatus, save,
      frame: async (viewport, options) => decodeFrame(await frame(viewport, options)),
    },
    save: async () => new Uint8Array(await save()),
    on: (name, listener) => client.on(name, listener),
    onFailure: (listener) => client.onFailure(listener),
    get failure() { return client.failure; },
    dispose: () => {
      if (peerSource) peerSource.disposed = true;
      return client.dispose();
    },
  };
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
  return session;
}

import {
  createSessionClient, requestWasmCompile, type SessionClient,
} from '../../../../shared/office-session/client';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure, type MethodPolicy, type Promisified } from '../../../../shared/office-session/types';
import { wasmAssetUrl } from '../wasm/asset';
import {
  createWorkbookPeerOpener, initWasm, resolveCollaborativeClientId,
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
const PEER_OPEN_UNITS = 256;
const PEER_OPEN_SLICE_MS = 8;

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
    clientId: resolveCollaborativeClientId(options),
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

function schedulePeerSlice<T>(slice: () => T): Promise<T> {
  const scheduler = (globalThis as typeof globalThis & {
    scheduler?: { postTask<T>(callback: () => T, options: { priority: 'background' }): Promise<T> };
  }).scheduler;
  if (scheduler?.postTask) return scheduler.postTask(slice, { priority: 'background' });
  return new Promise((resolve, reject) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.close();
      channel.port2.close();
      try { resolve(slice()); } catch (error) { reject(error); }
    };
    channel.port2.postMessage(undefined);
  });
}

function waitForPeerInput(source: WorkbookPeerSource): Promise<void> {
  return new Promise((resolve) => { source.wake = resolve; });
}

function stopPeerOpen(source: WorkbookPeerSource): void {
  source.disposed = true;
  source.bytes = undefined;
  source.hydration.length = 0;
  source.opener?.dispose();
  source.opener = undefined;
  source.wake?.();
  source.wake = undefined;
  void source.pending?.then((peer) => peer.dispose(), () => {});
}

async function openPeerFromSource(source: WorkbookPeerSource): Promise<WorkbookHandle> {
  let peer: WorkbookHandle | undefined;
  function check(): void {
    if (source.disposed) throw new SessionFailure('disposed', 'Workbook session was disposed');
    if (source.failure) throw source.failure;
  }
  try {
    while (!source.wasm && !source.module) {
      check();
      await waitForPeerInput(source);
    }
    check();
    await initWasm(source.wasm ?? source.module!);
    check();
    while (source.holds > 0) {
      await waitForPeerInput(source);
      check();
    }
    if (!source.bytes) throw new Error('Workbook peer source bytes are missing');
    const opener = createWorkbookPeerOpener(source.bytes, source.options);
    source.opener = opener;
    source.bytes = undefined;
    source.wasm = undefined;
    let ready = false;
    while (true) {
      check();
      if (source.holds > 0) {
        await waitForPeerInput(source);
        continue;
      }
      if (ready) break;
      const state = await schedulePeerSlice(() => {
        check();
        if (source.holds > 0) return;
        const deadline = performance.now() + PEER_OPEN_SLICE_MS;
        while (source.hydration.length > 0 && performance.now() < deadline) {
          opener.pushHydration(source.hydration.shift()!);
        }
        let state: number;
        do {
          check();
          state = opener.advance(PEER_OPEN_UNITS);
        } while (state !== 2 && state !== 1 && performance.now() < deadline);
        return state;
      });
      check();
      ready = state === 2;
      if (state === 1 && source.hydration.length === 0) {
        if (source.deliveryComplete) throw new WorkbookPeerHydrationError('missing-hydration',
          'Workbook worker hydration ended before the peer was ready');
        await waitForPeerInput(source);
      }
    }
    check();
    peer = opener.finish();
    if (peer.version() !== source.version) throw new WorkbookPeerHydrationError('version-mismatch',
      'Workbook source peer version differs from worker hydration');
    check();
    return peer;
  } catch (error) {
    peer?.dispose();
    if (source.disposed) throw new SessionFailure('disposed', 'Workbook session was disposed');
    const failure = error instanceof WorkbookPeerHydrationError ? error :
      new WorkbookPeerHydrationError('missing-hydration',
        `Workbook peer opening failed: ${error instanceof Error ? error.message : String(error)}`);
    source.failure = failure;
    throw failure;
  } finally {
    source.opener?.dispose();
    source.opener = undefined;
    source.bytes = undefined;
    source.hydration.length = 0;
    source.wake = undefined;
  }
}

/** @internal */
export function holdPeerOpen(session: WorkbookSession): () => void {
  const source = peerSources.get(session);
  if (!source || source.disposed) return () => {};
  source.holds += 1;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    source.holds -= 1;
    source.wake?.();
    source.wake = undefined;
  };
}

/** @internal */
export function hydratePeer(session: WorkbookSession): Promise<WorkbookHandle> {
  const source = peerSources.get(session);
  if (!source) return Promise.reject(new TypeError('Workbook session does not retain peer hydration state'));
  if (source.disposed) return Promise.reject(new SessionFailure('disposed', 'Workbook session was disposed'));
  return source.pending!;
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
        wasm: input.wasm instanceof ArrayBuffer ? input.wasm.slice(0) : undefined,
        hydration: [], receivedHydration: false, deliveryComplete: false, holds: 0, disposed: false,
      };
    }
    client = createSessionClient<Methods, Events>(transport, {
      methods: { ...WORKBOOK_SESSION_METHODS, ...WORKBOOK_INTERNAL_SESSION_METHODS },
      onWasmModuleMessage: ({ url, module, hydration, version, sequence }) => {
        if (url !== wasmAssetUrl().href) return;
        if (peerSource) {
          peerSource.module = module;
          if (hydration !== undefined) {
            if (typeof version !== 'string' || typeof sequence !== 'number' ||
              !Number.isSafeInteger(sequence) || sequence < 0) {
              const error = new WorkbookPeerHydrationError('missing-hydration',
                'Workbook worker hydration is missing its committed version or sequence');
              peerSource.failure = error;
              rejectHydration(error);
            } else if (peerSource.version !== undefined &&
              (peerSource.version !== version || peerSource.sequence !== sequence)) {
              const error = new WorkbookPeerHydrationError('version-mismatch',
                'Workbook worker hydration identity changed during opening');
              peerSource.failure = error;
              rejectHydration(error);
            } else {
              peerSource.hydration.push(hydration);
              peerSource.receivedHydration = true;
              peerSource.version = version;
              peerSource.sequence = sequence;
            }
          }
          peerSource.wake?.();
          peerSource.wake = undefined;
        }
        if (options.wasm === undefined && !options.worker && !wasmModules.has(url)) {
          wasmModules.set(url, module);
        }
      },
    });
    if (peerSource) {
      peerSource.pending = openPeerFromSource(peerSource);
      void peerSource.pending.catch(() => {});
    }
  } catch (error) {
    if (peerSource) stopPeerOpen(peerSource);
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
    if (peerSource) stopPeerOpen(peerSource);
    void client.dispose().catch(() => {});
    try { transport.close(); } catch {}
  };
  try {
    signal?.addEventListener('abort', abort, { once: true });
    if (signal?.aborted) abort();
    state = await Promise.race([
      client.callWithTransfer('open', [document, input], transfer), hydrationFailure,
    ]);
    if (peerSource) {
      peerSource.deliveryComplete = true;
      peerSource.wake?.();
      peerSource.wake = undefined;
    }
    if (signal?.aborted) throw new SessionFailure('disposed', 'Session was disposed');
    if (peerSource && !peerSource.module && !peerSource.wasm && !peerSource.opener) {
      throw new WorkbookPeerHydrationError('missing-module',
        'Workbook worker did not retain a compiled module for peer hydration');
    }
    if (peerSource && !peerSource.receivedHydration) {
      throw new WorkbookPeerHydrationError('missing-hydration',
        'Workbook worker did not retain peer calculation state');
    }
  } catch (error) {
    if (peerSource) stopPeerOpen(peerSource);
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
      if (peerSource) stopPeerOpen(peerSource);
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
      const list = reply.result.sheetInfo;
      state = {
        ...state, activeSheet: list.activeSheet,
        sheets: list.sheetIds.map((id, index) => ({ id, index, name: list.sheetNames[index] })),
      };
    }
    return reply;
  };
  const source = peerSource;
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

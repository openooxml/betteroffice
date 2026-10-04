import {
  createSessionClient, requestWasmCompile, type SessionClient,
} from '../../../../shared/office-session/client';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure, type MethodPolicy, type Promisified } from '../../../../shared/office-session/types';
import { wasmAssetUrl } from '../wasm/asset';
import type { OpenWorkbookOptions, Viewport } from '../wasm/loader';
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

type Events = { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] };
type Methods = WorkbookSessionMethods & WorkbookInternalSessionMethods;
const wasmModules = new Map<string, WebAssembly.Module>();

/**
 * Options for opening a workbook in a dedicated worker.
 * @experimental
 */
export interface OpenWorkbookSessionOptions extends OpenWorkbookOptions {
  worker?: () => Worker;
  wasm?: ArrayBuffer | WebAssembly.Module;
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
  document: ArrayBuffer; input: WorkbookSessionOpenOptions; transfer: Transferable[];
} {
  const document = copyBytes(bytes).buffer;
  const transfer: Transferable[] = [document];
  const input: WorkbookSessionOpenOptions = {
    collaborative: options.collaborative,
    clientId: options.clientId,
    calculation: options.calculation,
  };
  if (options.wasm instanceof ArrayBuffer ||
    Object.prototype.toString.call(options.wasm) === '[object ArrayBuffer]') {
    input.wasm = copyBytes(options.wasm as ArrayBuffer).buffer;
    transfer.push(input.wasm);
  } else if (options.wasm !== undefined) input.wasm = options.wasm;
  else if (!options.worker) input.wasm = wasmModules.get(wasmAssetUrl().href);
  return { document, input, transfer };
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
  try {
    ({ document, input, transfer } = prepareOpen(bytes, options));
    client = createSessionClient<Methods, Events>(transport, {
      methods: { ...WORKBOOK_SESSION_METHODS, ...WORKBOOK_INTERNAL_SESSION_METHODS },
      onWasmModule: options.wasm === undefined && !options.worker ? (url, module) => {
        if (url === wasmAssetUrl().href && !wasmModules.has(url)) wasmModules.set(url, module);
      } : undefined,
    });
  } catch (error) {
    try { transport.close(); } catch {}
    throw error;
  }
  let state: WorkbookSessionState;
  client.on('changed', (change) => { state = { ...state, ...change }; });
  client.onFailure(() => { state = { ...state, stage: 'failed' }; });
  const signal = options.signal;
  const abort = () => {
    void client.dispose().catch(() => {});
    try { transport.close(); } catch {}
  };
  try {
    signal?.addEventListener('abort', abort, { once: true });
    if (signal?.aborted) abort();
    state = await client.callWithTransfer('open', [document, input], transfer);
    if (signal?.aborted) throw new SessionFailure('disposed', 'Session was disposed');
  } catch (error) {
    await client.dispose();
    throw error;
  } finally { signal?.removeEventListener('abort', abort); }

  const calls = { ...client.call };
  for (const method of Object.keys(WORKBOOK_SESSION_METHODS) as (keyof WorkbookSessionMethods)[]) {
    Object.defineProperty(calls, method, { value: (...args: unknown[]) => {
      if (workbookSessionInternals.get(session)?.editPeerAttached) {
        const configured = WORKBOOK_SESSION_POLICIES[method];
        const policy = typeof configured === 'function'
          ? (configured as (...args: unknown[]) => MethodPolicy)(...args) : configured;
        if (policy.mutates) {
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
    dispose: () => client.dispose(),
  };
  const replay = async (envelope: WorkbookReplayEnvelope) => {
    const reply = await client.call.replay(envelope);
    if (envelope.op.method === 'setActiveSheet') {
      state = { ...state, activeSheet: envelope.op.args[0] };
    } else if (reply.result && 'sheetInfo' in reply.result) {
      const info = reply.result.sheetInfo;
      state = {
        ...state, activeSheet: info.activeSheet,
        sheets: info.sheetIds.map((id, index) => ({ id, index, name: info.sheetNames[index] })),
      };
    }
    return reply;
  };
  workbookSessionInternals.set(session, { replay, editPeerAttached: false });
  return session;
}

import {
  createSessionClient,
  createWorkerTransport,
  type Promisified,
  type SessionClient,
  type SessionFailure,
  type SessionTransport,
} from '../../../../shared/office-session';
import type { OpenWorkbookOptions, Viewport } from '../wasm/loader';
import {
  WORKBOOK_SESSION_METHODS,
  type WorkbookFrame,
  type WorkbookFrameOptions,
  type WorkbookSessionEvents,
  type WorkbookSessionMethods,
  type WorkbookSessionOpenOptions,
  type WorkbookSessionState,
  type WorkbookWireFrame,
} from './methods';

type Events = { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] };

/**
 * Options for opening a workbook in a dedicated worker.
 * @experimental
 */
export interface OpenWorkbookSessionOptions extends OpenWorkbookOptions {
  worker?: () => Worker;
  wasm?: ArrayBuffer | WebAssembly.Module;
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
  };
  if (options.wasm instanceof ArrayBuffer ||
    Object.prototype.toString.call(options.wasm) === '[object ArrayBuffer]') {
    input.wasm = copyBytes(options.wasm as ArrayBuffer).buffer;
    transfer.push(input.wasm);
  } else if (options.wasm !== undefined) input.wasm = options.wasm;
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
  let client: SessionClient<WorkbookSessionMethods, Events>;
  try {
    ({ document, input, transfer } = prepareOpen(bytes, options));
    client = createSessionClient<WorkbookSessionMethods, Events>(transport, {
      methods: WORKBOOK_SESSION_METHODS,
    });
  } catch (error) {
    try { transport.close(); } catch {}
    throw error;
  }
  let state: WorkbookSessionState;
  client.on('changed', (change) => { state = { ...state, ...change }; });
  client.onFailure(() => { state = { ...state, stage: 'failed' }; });
  try {
    state = await client.callWithTransfer('open', [document, input], transfer);
  } catch (error) {
    await client.dispose();
    throw error;
  }

  const {
    version, readCells, findText, validateEdits, applyEdits, frame, sheets, calculationStatus, save,
  } = client.call;
  return {
    get state() { return state; },
    call: {
      version, readCells, findText, validateEdits, applyEdits, sheets, calculationStatus, save,
      frame: async (viewport, options) => decodeFrame(await frame(viewport, options)),
    },
    save: async () => new Uint8Array(await save()),
    on: (name, listener) => client.on(name, listener),
    onFailure: (listener) => client.onFailure(listener),
    get failure() { return client.failure; },
    dispose: () => client.dispose(),
  };
}

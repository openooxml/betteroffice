import {
  createSessionClient, requestWasmCompile, type SessionClient,
} from '../../../../shared/office-session/client';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import type { Promisified, SessionFailure } from '../../../../shared/office-session/types';
import type { PptxFontFace } from '../types';
import { wasmAssetUrl } from '../wasm/asset';
import type { OpenPresentationOptions } from '../wasm/loader';
import { frameAssetIds } from './frame';
import {
  PRESENTATION_SESSION_METHODS,
  type PresentationFrame,
  type PresentationSessionEvents,
  type PresentationSessionFont,
  type PresentationSessionMethods,
  type PresentationSessionOpenOptions,
  type PresentationSessionState,
  type PresentationWireFrame,
} from './methods';

type Events = { [K in keyof PresentationSessionEvents]: PresentationSessionEvents[K] };
const wasmModules = new Map<string, WebAssembly.Module>();

/**
 * Options for opening a presentation in a dedicated worker.
 * @experimental
 */
export interface OpenPresentationSessionOptions extends OpenPresentationOptions {
  worker?: () => Worker;
  wasm?: ArrayBuffer | WebAssembly.Module;
}

/**
 * Async presentation access and its current local projection.
 * @experimental
 */
export interface PresentationSession {
  readonly state: PresentationSessionState;
  readonly call: Promisified<Omit<PresentationSessionMethods, 'open' | 'dispose' | 'frame'>> & {
    /**
     * @experimental A newer frame call replaces a queued one, which rejects with an error
     * named SessionSuperseded.
     */
    frame(slideIndex: number): Promise<PresentationFrame>;
  };
  save(): Promise<Uint8Array>;
  on<K extends keyof PresentationSessionEvents>(
    name: K, listener: (payload: PresentationSessionEvents[K]) => void
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

function decodeFrame(
  frame: PresentationWireFrame, cache: Map<string, Uint8Array>
): PresentationFrame {
  for (const { assetId, bytes } of frame.media) cache.set(assetId, new Uint8Array(bytes));
  const displayList = JSON.parse(new TextDecoder().decode(frame.displayList)) as
    PresentationFrame['displayList'];
  const media = new Map<string, Uint8Array>();
  for (const assetId of frameAssetIds(displayList)) {
    const bytes = cache.get(assetId);
    if (bytes !== undefined) media.set(assetId, bytes);
  }
  return { ...frame, displayList, media };
}

function copyFonts(
  faces: readonly PptxFontFace[] | undefined, transfer: Transferable[]
): PresentationSessionFont[] | undefined {
  return faces?.map((face) => {
    const bytes = copyBytes(face.bytes).buffer;
    transfer.push(bytes);
    return { family: face.family, bytes, bold: face.bold, italic: face.italic };
  });
}

function prepareOpen(bytes: Uint8Array | ArrayBuffer, options: OpenPresentationSessionOptions): {
  document: ArrayBuffer; input: PresentationSessionOpenOptions; transfer: Transferable[];
} {
  const document = copyBytes(bytes).buffer;
  const transfer: Transferable[] = [document];
  const input: PresentationSessionOpenOptions = {
    clientId: options.clientId,
    fonts: copyFonts(options.fonts, transfer),
    fallbackFonts: copyFonts(options.fallbackFonts, transfer),
  };
  if (options.initialUpdate !== undefined) {
    const update = copyBytes(options.initialUpdate);
    input.initialUpdate = update;
    transfer.push(update.buffer);
  }
  if (options.wasm instanceof ArrayBuffer) {
    input.wasm = options.wasm.slice(0);
    transfer.push(input.wasm);
  } else if (options.wasm !== undefined) input.wasm = options.wasm;
  else if (!options.worker) input.wasm = wasmModules.get(wasmAssetUrl().href);
  return { document, input, transfer };
}

/**
 * Opens a worker session, transferring copies so caller buffers stay usable.
 * @experimental
 */
export async function openPresentationSession(
  bytes: Uint8Array | ArrayBuffer,
  options: OpenPresentationSessionOptions = {}
): Promise<PresentationSession> {
  const transport = createWorkerTransport(
    options.worker ? options.worker() :
      new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' })
  );
  if (!options.worker && options.wasm === undefined && !wasmModules.has(wasmAssetUrl().href)) {
    requestWasmCompile(transport);
  }
  return createPresentationSession(bytes, options, transport);
}

/** Creates a presentation session over an internal transport. */
export async function createPresentationSession(
  bytes: Uint8Array | ArrayBuffer,
  options: OpenPresentationSessionOptions,
  transport: SessionTransport
): Promise<PresentationSession> {
  let document: ArrayBuffer;
  let input: PresentationSessionOpenOptions;
  let transfer: Transferable[];
  let client: SessionClient<PresentationSessionMethods, Events>;
  try {
    ({ document, input, transfer } = prepareOpen(bytes, options));
    client = createSessionClient<PresentationSessionMethods, Events>(transport, {
      methods: PRESENTATION_SESSION_METHODS,
      onWasmModule: options.wasm === undefined && !options.worker ? (url, module) => {
        if (url === wasmAssetUrl().href && !wasmModules.has(url)) wasmModules.set(url, module);
      } : undefined,
    });
  } catch (error) {
    try { transport.close(); } catch {}
    throw error;
  }
  let state: PresentationSessionState;
  client.on('changed', (change) => { state = { ...state, ...change }; });
  client.onFailure(() => { state = { ...state, stage: 'failed' }; });
  try {
    state = await client.callWithTransfer('open', [document, input], transfer);
  } catch (error) {
    await client.dispose();
    throw error;
  }

  const media = new Map<string, Uint8Array>();
  const {
    version, readContent, findText, validateEdits, applyEdits, frame, slides, slideSize, save,
  } = client.call;
  return {
    get state() { return state; },
    call: {
      version, readContent, findText, validateEdits, applyEdits, slides, slideSize, save,
      frame: async (slideIndex) => decodeFrame(await frame(slideIndex), media),
    },
    save: async () => new Uint8Array(await save()),
    on: (name, listener) => client.on(name, listener),
    onFailure: (listener) => client.onFailure(listener),
    get failure() { return client.failure; },
    dispose: async () => {
      try { await client.dispose(); } finally { media.clear(); }
    },
  };
}

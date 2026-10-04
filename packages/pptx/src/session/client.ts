import {
  createSessionClient,
  createWorkerTransport,
  type Promisified,
  type SessionFailure,
  type SessionTransport,
} from '../../../../shared/office-session';
import type { PptxFontFace } from '../types';
import type { OpenPresentationOptions } from '../wasm/loader';
import {
  PRESENTATION_SESSION_METHODS,
  type PresentationSessionEvents,
  type PresentationSessionFont,
  type PresentationSessionMethods,
  type PresentationSessionOpenOptions,
  type PresentationSessionState,
} from './methods';

type Events = { [K in keyof PresentationSessionEvents]: PresentationSessionEvents[K] };

/** Options for opening a presentation in a dedicated worker. */
export interface OpenPresentationSessionOptions extends OpenPresentationOptions {
  worker?: () => Worker;
  wasm?: ArrayBuffer | WebAssembly.Module;
  /** @internal */
  transport?: SessionTransport;
}

/** Async presentation access and its current local projection. */
export interface PresentationSession {
  readonly state: PresentationSessionState;
  readonly call: Promisified<Omit<PresentationSessionMethods, 'open' | 'dispose'>>;
  save(): Promise<Uint8Array>;
  on<K extends keyof PresentationSessionEvents>(
    name: K, listener: (payload: PresentationSessionEvents[K]) => void
  ): () => void;
  onFailure(listener: (failure: SessionFailure) => void): () => void;
  readonly failure: SessionFailure | undefined;
  dispose(): Promise<void>;
}

function copyFonts(
  faces: readonly PptxFontFace[] | undefined, transfer: Transferable[]
): PresentationSessionFont[] | undefined {
  return faces?.map((face) => {
    const bytes = new Uint8Array(face.bytes).buffer;
    transfer.push(bytes);
    return { ...face, bytes };
  });
}

/** Opens a worker session, transferring copies so caller buffers stay usable. */
export async function openPresentationSession(
  bytes: Uint8Array | ArrayBuffer,
  options: OpenPresentationSessionOptions = {}
): Promise<PresentationSession> {
  const document = bytes instanceof ArrayBuffer ? bytes.slice(0) : new Uint8Array(bytes).buffer;
  const transfer: Transferable[] = [document];
  const input: PresentationSessionOpenOptions = {
    clientId: options.clientId,
    fonts: copyFonts(options.fonts, transfer),
    fallbackFonts: copyFonts(options.fallbackFonts, transfer),
  };
  if (options.initialUpdate !== undefined) {
    const update = new Uint8Array(options.initialUpdate);
    input.initialUpdate = update;
    transfer.push(update.buffer);
  }
  if (options.wasm instanceof ArrayBuffer) {
    input.wasm = options.wasm.slice(0);
    transfer.push(input.wasm);
  } else if (options.wasm !== undefined) input.wasm = options.wasm;

  const transport = options.transport ?? createWorkerTransport(
    options.worker ? options.worker() :
      new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' })
  );
  const client = createSessionClient<PresentationSessionMethods, Events>(transport, {
    methods: PRESENTATION_SESSION_METHODS,
  });
  let state: PresentationSessionState;
  client.on('changed', (change) => { state = { ...state, ...change }; });
  try {
    state = await client.callWithTransfer('open', [document, input], transfer);
  } catch (error) {
    await client.dispose();
    throw error;
  }

  const { version, readContent, findText, validateEdits, applyEdits, slides, slideSize, save } = client.call;
  return {
    get state() { return state; },
    call: { version, readContent, findText, validateEdits, applyEdits, slides, slideSize, save },
    save: async () => new Uint8Array(await save()),
    on: (name, listener) => client.on(name, listener),
    onFailure: (listener) => client.onFailure(listener),
    get failure() { return client.failure; },
    dispose: () => client.dispose(),
  };
}

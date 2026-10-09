import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import {
  adoptPresentationPeerIdentity, initWasm, openPresentationPeerDeck,
  registerPresentationPeerFonts, type PresentationPeerHandle,
} from '../wasm/loader';
import { createPresentationEditorClient, type OpenPresentationSessionOptions } from './client';
import {
  createPresentationEditPeer, type PresentationEditPeer, type PptxWorkerEditorAccess,
} from './editPeer';
import { presentationEditPeerInternals } from './editPeerInternals';
import { peerHydrationError } from './editorPeerHydrationError';
import type { PresentationSessionState } from './methods';
import {
  PptxPeerHydrationError, PptxPeerNotReadyError, PptxWorkerEditorCollaborationError,
  PptxWorkerEditorDisposedError, PptxWorkerEditorFailedError,
} from './peerHydrationError';
import {
  presentationEditorSessionInternals,
  type PptxWorkerEditorFrame, type PptxWorkerEditorOperation, type PptxWorkerEditorReply,
} from './replay';

export type PptxWorkerEditorStage = 'opening' | 'hydrating' | 'ready' | 'failed' | 'recovering' | 'disposed';
export interface PptxWorkerEditorState {
  readonly stage: PptxWorkerEditorStage;
  readonly sequence: number;
  readonly acknowledgedSequence: number;
  readonly version?: string;
  readonly projection?: PresentationSessionState;
  readonly initialFrame?: PptxWorkerEditorFrame;
  readonly failure?: PptxWorkerEditorFailedError;
}
export interface PptxWorkerEditorOptions extends OpenPresentationSessionOptions {
  collaboration?: unknown;
  signal?: AbortSignal;
  onError?(error: PptxWorkerEditorFailedError): void;
}
export interface PptxWorkerEditorRecovery {
  bytes: Uint8Array;
  recovery: true;
}
export interface PptxWorkerEditorSession {
  readonly state: PptxWorkerEditorState;
  readonly hydrated: boolean;
  readonly failure: PptxWorkerEditorFailedError | undefined;
  subscribe(listener: (state: PptxWorkerEditorState) => void): () => void;
  whenHydrated(): Promise<void>;
  handleAsync(): Promise<PptxWorkerEditorAccess>;
  apply(op: PptxWorkerEditorOperation): PptxWorkerEditorReply;
  flush(): Promise<void>;
  saveAsync(): Promise<Uint8Array>;
  frame(slideId: string): Promise<PptxWorkerEditorFrame>;
  recoverySave(): Promise<PptxWorkerEditorRecovery>;
  dispose(): Promise<void>;
}

export function createPptxWorkerEditorSession(
  bytes: Uint8Array | ArrayBuffer, options: PptxWorkerEditorOptions = {}
): PptxWorkerEditorSession {
  preflight(options);
  const transport = createWorkerTransport(options.worker ? options.worker() :
    new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' }));
  return createPresentationEditorSession(bytes, options, transport);
}

function preflight(options: PptxWorkerEditorOptions): void {
  if (options.collaboration != null) {
    throw new PptxWorkerEditorCollaborationError();
  }
  if (options.signal?.aborted) throw new PptxWorkerEditorDisposedError();
}

function ownFonts(faces: PptxWorkerEditorOptions['fonts']) {
  return faces?.map((face) => ({ ...face, bytes: new Uint8Array(face.bytes).slice() }));
}

export function createPresentationEditorSession(
  bytes: Uint8Array | ArrayBuffer, options: PptxWorkerEditorOptions, transport: SessionTransport
): PptxWorkerEditorSession {
  preflight(options);
  let connection: ReturnType<typeof createPresentationEditorClient>;
  let source: Uint8Array;
  let peerOptions: Pick<PptxWorkerEditorOptions, 'fonts' | 'fallbackFonts' | 'initialUpdate'>;
  try {
    source = ArrayBuffer.isView(bytes)
      ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength).slice()
      : new Uint8Array(bytes).slice();
    peerOptions = { fonts: ownFonts(options.fonts), fallbackFonts: ownFonts(options.fallbackFonts),
      initialUpdate: options.initialUpdate?.slice() };
    connection = createPresentationEditorClient(source, { ...options, ...peerOptions }, transport);
  } catch (cause) {
    try { transport.close(); } catch {}
    throw cause;
  }
  let state: PptxWorkerEditorState = { stage: 'opening', sequence: 0, acknowledgedSequence: 0 };
  let peer: PresentationPeerHandle | undefined;
  let incomplete: PresentationPeerHandle | undefined;
  let edits: PresentationEditPeer | undefined;
  let recovery: Uint8Array | undefined;
  let disposal: Promise<void> | undefined;
  const listeners = new Set<(state: PptxWorkerEditorState) => void>();
  let rejectEnded!: (error: Error) => void;
  const ended = new Promise<never>((_, reject) => { rejectEnded = reject; });
  void ended.catch(() => {});

  function publish(change: Partial<PptxWorkerEditorState>): void {
    state = Object.freeze({ ...state, ...change });
    for (const listener of [...listeners]) {
      if (listeners.has(listener)) { try { listener(state); } catch {} }
    }
  }
  function current(): void {
    if (state.stage === 'disposed') throw new PptxWorkerEditorDisposedError();
    if (state.failure) throw state.failure;
  }
  function ready(): PresentationEditPeer {
    current();
    if (state.stage !== 'ready' || !edits) throw new PptxPeerNotReadyError();
    return edits;
  }
  function disposeIncomplete(): void {
    const opened = incomplete;
    incomplete = undefined;
    if (peer === opened) peer = undefined;
    try { opened?.dispose(); } catch {}
  }
  function fail(cause: unknown): void {
    if (state.failure || state.stage === 'disposed') return;
    const error = cause instanceof PptxWorkerEditorFailedError ? cause : new PptxWorkerEditorFailedError(cause);
    disposeIncomplete();
    publish({ stage: 'failed', failure: error });
    rejectEnded(error);
    if (edits) presentationEditPeerInternals.get(edits)?.fail(error);
    try { options.onError?.(error); } catch {}
    void connection.dispose();
  }
  function preserve(): Uint8Array {
    if (!recovery) {
      if (!peer) throw new PptxPeerNotReadyError();
      recovery = new Uint8Array(peer.save()).slice();
    }
    if (state.stage === 'disposed' && peer) {
      peer.dispose();
      peer = undefined;
    }
    return recovery.slice();
  }
  const { client } = presentationEditorSessionInternals.get(connection)!;
  const offFailure = connection.onFailure(fail);
  const yieldStage = () => new Promise<void>((resolve) => setTimeout(resolve, 0));
  const hydration = (async () => {
    try {
      const baseline = await Promise.race([connection.begin(), ended]);
      current();
      if (baseline.sequence !== 0) throw new PptxPeerHydrationError('sequence', 'Editor baseline is not at zero');
      publish({ stage: 'hydrating', projection: baseline.state, version: baseline.version });
      current();
      await initWasm(baseline.module);
      await yieldStage();
      current();
      incomplete = openPresentationPeerDeck(source, baseline.hydration, peerOptions);
      await yieldStage();
      current();
      await registerPresentationPeerFonts(incomplete);
      await yieldStage();
      current();
      adoptPresentationPeerIdentity(incomplete);
      if (incomplete.version() !== baseline.version) {
        throw new PptxPeerHydrationError('version', 'Hydrated peer version differs');
      }
      const attached = await Promise.race([client.call.attachPeer(incomplete.version(), 0), ended]);
      current();
      if (attached.sequence !== 0 || attached.version !== incomplete.version()) {
        throw new PptxPeerHydrationError('attachment', 'Worker attachment acknowledgement differs');
      }
      const slideId = baseline.state.slides[0]?.id;
      const frame = slideId === undefined ? undefined : await Promise.race([connection.frame(slideId), ended]);
      current();
      if (frame && (frame.sequence !== 0 || frame.version !== baseline.version || frame.slideId !== slideId)) {
        throw new PptxPeerHydrationError('frame', 'Initial worker frame differs');
      }
      peer = incomplete;
      edits = createPresentationEditPeer({
        peer, session: connection, guard: () => { ready(); }, onFailure: fail,
        onProgress: (sequence, acknowledgedSequence, projection) => {
          if (state.stage !== 'ready') return;
          publish({ sequence, acknowledgedSequence, version: edits?.version,
            ...(projection ? { projection } : {}) });
        },
      });
      incomplete = undefined;
      publish({ stage: 'ready', initialFrame: frame });
      current();
    } catch (cause) {
      const error = peerHydrationError(cause);
      fail(error);
      throw state.stage === 'disposed' ? new PptxWorkerEditorDisposedError() : state.failure ?? error;
    } finally {
      disposeIncomplete();
    }
  })();
  void hydration.catch(() => {});

  const owner: PptxWorkerEditorSession = {
    get state() { return state; },
    get hydrated() { return state.stage === 'ready'; },
    get failure() { return state.failure; },
    subscribe(listener) {
      if (state.stage === 'disposed') throw new PptxWorkerEditorDisposedError();
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    whenHydrated: async () => { current(); await Promise.race([hydration, ended]); current(); },
    handleAsync: async () => { await owner.whenHydrated(); return ready().access; },
    apply: (op) => ready().apply(op),
    flush: async () => ready().flush(),
    saveAsync: async () => ready().saveAsync(),
    frame: async (slideId) => {
      current();
      if (!state.projection) throw new PptxPeerNotReadyError();
      try {
        const frame = await Promise.race([connection.frame(slideId), ended]);
        current();
        return frame;
      } catch (cause) {
        current();
        if (cause instanceof Error && cause.name === 'RangeError') throw cause;
        fail(cause);
        current();
        throw cause;
      }
    },
    recoverySave: async () => {
      if (state.stage === 'recovering') throw new PptxPeerHydrationError('stage', 'Recovery is already in progress');
      if (state.stage !== 'disposed' && !state.failure) throw new PptxPeerNotReadyError();
      const retired = state.stage === 'disposed';
      if (!retired) publish({ stage: 'recovering' });
      try { return { bytes: preserve(), recovery: true }; }
      finally { if (state.stage !== 'disposed') publish({ stage: 'failed' }); }
    },
    dispose() {
      if (disposal) return disposal;
      let resolveDisposal!: () => void;
      let rejectDisposal!: (cause: unknown) => void;
      disposal = new Promise<void>((resolve, reject) => { resolveDisposal = resolve; rejectDisposal = reject; });
      publish({ stage: 'disposed' });
      disposeIncomplete();
      rejectEnded(new PptxWorkerEditorDisposedError());
      edits?.dispose();
      offFailure();
      options.signal?.removeEventListener('abort', abort);
      try { if (peer) preserve(); } catch {}
      listeners.clear();
      void connection.dispose().then(resolveDisposal, rejectDisposal);
      return disposal;
    },
  };
  const abort = () => { void owner.dispose(); };
  options.signal?.addEventListener('abort', abort, { once: true });
  if (options.signal?.aborted) abort();
  return owner;
}

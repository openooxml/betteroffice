import { presentationCollaborationPipes } from './commandCollaboration';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { SessionFailure } from '../../../../shared/office-session/types';
import { requestWasmCompile } from '../../../../shared/office-session/client';
import { wasmAssetUrl } from '../wasm/asset';
import type { OpenPresentationSessionOptions } from './client';
import { createSharedPresentationCommandClient } from './commandClient';
import { presentationSharedInputs } from './commandSharedInput';
import { CommandOutcomeUnknown } from './commandProtocol';
import type {
  PresentationCommandClient, PresentationCommandFailure, PresentationCommandOpenResult,
} from './commandProtocol';
import {
  PptxWorkerEditorDisposedError, PptxWorkerEditorFailedError,
} from './peerHydrationError';

const wasmModules = new Map<string, WebAssembly.Module>();

/** @internal */
export interface PptxWorkerSessionOptions extends OpenPresentationSessionOptions {
  readOnly?: boolean;
  observeChanges?: boolean;
  activeSlideIndex?: number;
  collaboration?: { clientId: number; initialUpdate?: Uint8Array };
  signal?: AbortSignal;
  onError?(error: PptxWorkerEditorFailedError): void;
}

/** @internal */
export interface PptxWorkerSession {
  readonly ready: Promise<PresentationCommandClient>;
  readonly failure: PresentationCommandFailure | undefined;
  subscribeFailure(listener: (error: PresentationCommandFailure) => void): () => void;
  dispose(): Promise<void>;
}

/** @internal */
export class PptxWorkerUnsupportedError extends Error {
  readonly code = 'worker-api-unavailable';

  constructor(readonly operation: string, readonly alternative?: string) {
    super(`${operation} is unavailable in the worker presentation editor${alternative ? `; use ${alternative} instead` : ''}`);
    this.name = 'PptxWorkerUnsupportedError';
  }
}

/** @internal */
export function createPptxWorkerSession(
  bytes: Uint8Array | ArrayBuffer, options: PptxWorkerSessionOptions = {},
): PptxWorkerSession {
  if (options.signal?.aborted) throw new PptxWorkerEditorDisposedError();
  const cacheUrl = !options.worker && options.wasm === undefined ? wasmAssetUrl().href : undefined;
  const wasm = options.wasm ?? (cacheUrl ? wasmModules.get(cacheUrl) : undefined);
  const transport = createWorkerTransport(options.worker ? options.worker() :
    new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' }));
  if (cacheUrl && !wasm) requestWasmCompile(transport);
  const input = {
    session: crypto.randomUUID(), incarnation: crypto.randomUUID(),
    clientId: options.collaboration?.clientId ?? options.clientId, collaboration: options.collaboration != null,
    fonts: options.fonts, fallbackFonts: options.fallbackFonts,
    initialUpdate: options.collaboration?.initialUpdate ?? options.initialUpdate, wasm,
    observeChanges: options.observeChanges, readOnly: options.readOnly, activeSlideIndex: options.activeSlideIndex,
  };
  const listeners = new Set<(error: PresentationCommandFailure) => void>();
  const connectionErrors = new Set<(error: unknown) => void>();
  let failure: PresentationCommandFailure | undefined;
  let disposed = false;
  let disposal: Promise<void> | undefined;
  let client: PresentationCommandClient | undefined;
  let offClientFailure: (() => void) | undefined;
  let failedState: PresentationCommandOpenResult | undefined;
  let rejectEnded!: (error: Error) => void;
  const ended = new Promise<never>((_, reject) => { rejectEnded = reject; });
  void ended.catch(() => {});

  function fail(cause: unknown): void {
    if (disposed || failure) return;
    failure = cause instanceof SessionFailure || cause instanceof CommandOutcomeUnknown ? cause
      : new SessionFailure(cause instanceof Error && cause.name === 'RuntimeError' ? 'trap' : 'crash',
        cause instanceof Error ? cause.message : String(cause));
    if (client) failedState = { ...client.state, editState: { ...client.state.editState, readOnly: true,
      toolbar: { ...client.state.editState.toolbar, canEditText: false } } };
    rejectEnded(failure);
    for (const listener of [...listeners]) {
      try { listener(failure); } catch {}
    }
    try { options.onError?.(new PptxWorkerEditorFailedError(failure)); } catch {}
    for (const listener of [...connectionErrors]) listener(failure);
    transport.close();
  }

  function current(): void {
    if (disposed) throw new PptxWorkerEditorDisposedError();
    if (failure) throw failure;
  }

  async function perform<T>(operation: () => Promise<T>): Promise<T> {
    current();
    try {
      const value = await operation();
      current();
      return value;
    } catch (error) {
      if (error instanceof SessionFailure || error instanceof Error &&
        (error.name === 'CommandOutcomeUnknown' || error.name === 'RuntimeError')) fail(error);
      throw failure ?? error;
    }
  }

  const offError = transport.onError(fail);
  const monitored: SessionTransport = {
    post: (message, transfer) => transport.post(message, transfer),
    listen: (listener) => transport.listen((message) => {
      if (isHostMessage(message)) {
        if (message.kind === 'wasm-module' && cacheUrl && !disposed && !failure && message.url === cacheUrl && !wasmModules.has(cacheUrl)) {
          wasmModules.set(cacheUrl, message.module);
        }
        if (message.kind === 'failure') fail(new SessionFailure(message.code, message.message, message.diagnostics));
      }
      listener(message);
    }),
    onError(listener) {
      connectionErrors.add(listener);
      const off = transport.onError(listener);
      return () => { connectionErrors.delete(listener); off(); };
    },
    close() {
      if (!disposed && !failure) fail(new SessionFailure('crash', 'Presentation command connection closed'));
      transport.close();
    },
  };
  const opening = createSharedPresentationCommandClient(bytes, input, monitored).then((opened) => {
    client = opened;
    offClientFailure = opened.subscribeFailure(fail);
    if (disposed) offClientFailure();
    current();
    const facade = {
      get state() { return failedState ?? opened.state; },
      dispatch: (intent, basis) => perform(() => opened.dispatch(intent, basis)),
      flush: (seq) => perform(() => opened.flush(seq)),
      read: (request) => perform(() => opened.read(request)),
      save: () => perform(() => opened.save()),
      access: (method, ...args) => perform(() => opened.access!(method, ...args)),
      checkpoint: () => perform(() => opened.checkpoint!()),
      preparePicture: (bytes) => perform(() => opened.preparePicture!(bytes)),
      thumbnail: (slideId, knownRevision) => perform(() => opened.thumbnail(slideId, knownRevision)),
      subscribe: (listener) => opened.subscribe(listener),
      subscribeFrame: (listener) => opened.subscribeFrame(listener),
      subscribeMedia: (listener) => opened.subscribeMedia(listener),
      subscribeUpdate: (listener) => opened.subscribeUpdate!(listener),
      subscribeFailure: (listener) => session.subscribeFailure(listener),
      dispose: () => session.dispose(),
    } satisfies PresentationCommandClient;
    const pipe = presentationCollaborationPipes.get(opened);
    if (pipe) presentationCollaborationPipes.set(facade, pipe);
    const sharedInput = presentationSharedInputs.get(opened);
    if (sharedInput) presentationSharedInputs.set(facade, {
      subscribeSent: (listener) => sharedInput.subscribeSent(listener),
      submit(intent, basis, after) {
        current();
        const submission = sharedInput.submit(intent, basis, after);
        return { ticket: submission.ticket, operation: perform(() => submission.operation) };
      },
    });
    return facade;
  }).catch((error: unknown) => { if (!disposed) fail(error); throw failure ?? error; });
  const ready = Promise.race([opening, ended]);
  void ready.catch(() => {});
  const session: PptxWorkerSession = {
    ready,
    get failure() { return failure; },
    subscribeFailure(listener) {
      if (failure) listener(failure);
      else if (!disposed) listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    dispose() {
      if (disposal) return disposal;
      disposed = true;
      rejectEnded(new PptxWorkerEditorDisposedError());
      options.signal?.removeEventListener('abort', abort);
      offError();
      offClientFailure?.();
      listeners.clear();
      for (const listener of [...connectionErrors]) listener(new SessionFailure('disposed', 'Session was disposed'));
      transport.close();
      disposal = client ? client.dispose() : opening.then(
        async () => { await client?.dispose(); }, async () => { await client?.dispose(); },
      );
      return disposal;
    },
  };
  function abort(): void { void session.dispose(); }
  options.signal?.addEventListener('abort', abort, { once: true });
  if (options.signal?.aborted) abort();
  return session;
}

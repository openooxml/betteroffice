import type { PresentationEditorConnection } from './client';
import type { DeckSnapshot, HitTestResult } from '../types';
import {
  PresentationPeerError, replayPresentation,
  type PresentationHandle, type PresentationPeerHandle, type PresentationReplayEnvelope,
} from '../wasm/loader';
import { presentationEditPeerInternals } from './editPeerInternals';
import { peerHydrationError } from './editorPeerHydrationError';
import {
  PptxPeerHydrationError, PptxWorkerEditorDisposedError, PptxWorkerEditorFailedError,
} from './peerHydrationError';
import {
  matchingReplayReply, PRESENTATION_REPLAY_MUTATORS, presentationEditorSessionInternals,
  type PptxWorkerEditorOperation, type PptxWorkerEditorReply,
} from './replay';
import type { PresentationSessionState } from './methods';

const readMethods = [
  'snapshot', 'story', 'anchorCaret', 'resolveCaretAnchor', 'version', 'readContent', 'findText',
  'validateEdits', 'searchText', 'isProposalsAvailable', 'listProposals', 'layoutSlide', 'hitTest',
  'mediaBytes', 'comments', 'canUndo', 'canRedo', 'undoCaptureMode', 'onUpdate',
] as const satisfies readonly (keyof PresentationHandle)[];

export interface PptxWorkerEditorInteractionAccess {
  snapshot(): { snapshot: DeckSnapshot; keys: Record<string, string> };
  key(index: number): string;
  activate(slideId: string, key: string): boolean;
  hitTest(slideId: string, x: number, y: number): HitTestResult | null;
}
export type PptxWorkerEditorAccess = Pick<PresentationHandle,
  typeof readMethods[number] | PptxWorkerEditorOperation['method'] | 'clientId'> & {
    readonly interaction: PptxWorkerEditorInteractionAccess;
  };

export interface PresentationEditPeer {
  readonly access: PptxWorkerEditorAccess;
  readonly sequence: number;
  readonly acknowledgedSequence: number;
  readonly version: string;
  apply(op: PptxWorkerEditorOperation): PptxWorkerEditorReply;
  flush(): Promise<void>;
  saveAsync(): Promise<Uint8Array>;
  dispose(): void;
}

function ownOperation(op: PptxWorkerEditorOperation): PptxWorkerEditorOperation {
  const owned = structuredClone(op);
  const seen = new WeakSet<object>();
  function check(value: unknown): void {
    if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new PptxPeerHydrationError('arguments', 'Editor arguments must be finite');
    }
    if (value !== null && typeof value === 'object') {
      if (seen.has(value)) throw new PptxPeerHydrationError('arguments', 'Editor arguments must be acyclic');
      seen.add(value);
      for (const child of Object.values(value)) check(child);
      seen.delete(value);
      Object.freeze(value);
    }
  }
  check(owned);
  return owned;
}

export function createPresentationEditPeer(options: {
  peer: PresentationPeerHandle;
  session: PresentationEditorConnection;
  guard(): void;
  onFailure(error: PptxWorkerEditorFailedError): void;
  onProgress(sequence: number, acknowledgedSequence: number, projection?: PresentationSessionState): void;
}): PresentationEditPeer {
  const { peer, session } = options;
  const internal = presentationEditorSessionInternals.get(session);
  if (!internal) throw new TypeError('Presentation session does not support editor replay');
  if (internal.peerAttached) throw new PptxPeerHydrationError('attachment', 'Presentation peer is already attached');
  const { client } = internal;
  let sequence = 0;
  let acknowledgedSequence = 0;
  let currentVersion = peer.version();
  let failure: PptxWorkerEditorFailedError | undefined;
  let disposed = false;
  let applying = false;
  let tail = Promise.resolve();
  let rejectFailure!: (error: Error) => void;
  const failed = new Promise<never>((_, reject) => { rejectFailure = reject; });
  void failed.catch(() => {});

  function fail(cause: unknown): void {
    if (failure || disposed) return;
    failure = cause instanceof PptxWorkerEditorFailedError ? cause
      : new PptxWorkerEditorFailedError(peerHydrationError(cause));
    rejectFailure(failure);
    options.onFailure(failure);
  }
  function guard(): void {
    options.guard();
    if (disposed) throw new PptxWorkerEditorDisposedError();
    if (client.failure) fail(client.failure);
    if (failure) throw failure;
  }
  function enqueue<T>(operation: () => Promise<T>): Promise<T> {
    const queued = tail.then(async () => {
      guard();
      const value = await Promise.race([operation(), failed]);
      guard();
      return value;
    }).catch((cause: unknown) => {
      fail(cause);
      throw failure ?? cause;
    });
    tail = queued.then(() => {}, () => {});
    return Promise.race([queued, failed]);
  }
  function apply(op: PptxWorkerEditorOperation): PptxWorkerEditorReply {
    guard();
    if (applying) throw new PptxPeerHydrationError('reentrant', 'Presentation editor mutation cannot reenter');
    applying = true;
    let captured = false;
    let replayStarted = false;
    try {
      const envelope: PresentationReplayEnvelope = {
        sequence: sequence + 1, baseVersion: currentVersion, op: ownOperation(op),
      };
      replayStarted = true;
      return replayPresentation(peer, envelope, (reply) => {
        captured = true;
        if (!reply.consumed) return;
        sequence = reply.sequence;
        currentVersion = reply.version;
        const consumed = Object.freeze({ ...envelope, expectedOutcome: reply.outcome });
        void enqueue(async () => {
          const remote = await client.call.replay(consumed);
          if (!matchingReplayReply(reply, remote)) {
            throw new PptxPeerHydrationError('ack-mismatch', 'Presentation replay acknowledgement differs');
          }
          acknowledgedSequence = remote.sequence;
          options.onProgress(sequence, acknowledgedSequence, remote.projection);
        }).catch(() => {});
        options.onProgress(sequence, acknowledgedSequence);
      });
    } catch (cause) {
      if (captured || replayStarted && !(cause instanceof PresentationPeerError &&
        ['arguments', 'invalidJson', 'method'].includes(cause.code))) fail(cause);
      throw failure ?? cause;
    } finally {
      applying = false;
    }
  }
  function fence<T>(operation: (watermark: number, version: string) => Promise<T>): Promise<T> {
    try { guard(); } catch (cause) { return Promise.reject(cause); }
    const watermark = sequence;
    const version = currentVersion;
    return enqueue(async () => {
      const position = await client.call.flush(watermark, version);
      if (position.sequence !== watermark || position.version !== version || acknowledgedSequence !== watermark) {
        throw new PptxPeerHydrationError('ack-mismatch', 'Presentation flush acknowledgement differs');
      }
      return operation(watermark, version);
    });
  }

  const access = {} as PptxWorkerEditorAccess;
  function guardedRead(target: object, method: string) {
    return (...args: unknown[]) => {
      guard();
      return (target as Record<string, (...args: unknown[]) => unknown>)[method](...args);
    };
  }
  Object.defineProperty(access, 'clientId', { enumerable: true, get: () => { guard(); return peer.clientId; } });
  for (const method of readMethods) {
    Object.defineProperty(access, method, { enumerable: true, value: guardedRead(peer, method) });
  }
  const cacheSymbol = Symbol.for('@betteroffice/pptx/slide-layout-cache');
  const cache = (peer as unknown as Record<symbol, PptxWorkerEditorInteractionAccess>)[cacheSymbol];
  const interaction = Object.freeze(Object.fromEntries(['snapshot', 'key', 'activate', 'hitTest']
    .map((method) => [method, guardedRead(cache, method)]))) as unknown as PptxWorkerEditorInteractionAccess;
  Object.defineProperty(access, 'interaction', { enumerable: true, value: interaction });
  Object.defineProperty(access, cacheSymbol, { value: interaction });
  for (const method of Object.keys(PRESENTATION_REPLAY_MUTATORS) as PptxWorkerEditorOperation['method'][]) {
    Object.defineProperty(access, method, { enumerable: true, value: (...args: unknown[]) =>
      apply({ method, args } as PptxWorkerEditorOperation).outcome.result });
  }
  const offFailure = client.onFailure(fail);
  const edits: PresentationEditPeer = {
    access: Object.freeze(access),
    get sequence() { return sequence; },
    get acknowledgedSequence() { return acknowledgedSequence; },
    get version() { return currentVersion; },
    apply,
    flush: () => fence(async () => {}),
    saveAsync: () => fence(async (watermark, version) =>
      new Uint8Array(await client.call.editorSave(watermark, version))),
    dispose() {
      if (disposed) return;
      disposed = true;
      rejectFailure(new PptxWorkerEditorDisposedError());
      offFailure();
    },
  };
  internal.peerAttached = true;
  presentationEditPeerInternals.set(edits, {
    fail,
    whenAcknowledged: () => {
      try { guard(); } catch (cause) { return Promise.reject(cause); }
      return Promise.race([tail, failed]);
    },
  });
  return edits;
}

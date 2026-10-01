import type {
  DocxParagraphAnchor,
  DocxParagraphAnchorResult,
  DocxParagraphIdentitySnapshot,
  DocxProposalRequest,
  DocxProposalResult,
  DocxProposalSnapshot,
  DocxProposalStateRequest,
  DocxProposalWithdrawRequest,
  DocxReadParagraphsRequest,
  DocxReadParagraphsResult,
  ProposalGeometryMirror,
  ResidentProposalReply,
  YrsSession,
  resolveNavigationTarget,
} from '@betteroffice/docx/yrs';
import type { WorkerOpenedDocument } from '../hooks/useDisplayList';
import { proposalRevisionPreview, resolveMirroredNavigationTarget } from '@betteroffice/docx/yrs';
import {
  awaitWorkerOpenReplica,
  workerOpenReplicaPending,
  workerOpenReplicaStarted,
} from './workerOpenReplica';

export interface WorkerProposalAuthority {
  /** The session mirrors the worker registry and version. */
  readonly initialized: boolean;
  restart(): void;
  /** Initializes once; rejects when the worker cannot answer. */
  initialize(): Promise<void>;
  /** Mirrored geometry until hand-over. */
  geometry(): ProposalGeometryMirror | null;
  /** Reseeding would lose worker changes, including those of a state change still in flight. */
  holdsWorkerState(): boolean;
  /** Reseeding would lose worker changes the main thread has already observed. */
  holdsCommittedWorkerState(): boolean;
  propose(
    request: DocxProposalRequest,
    main: (request: DocxProposalRequest) => Promise<DocxProposalResult>
  ): Promise<DocxProposalResult>;
  setStates(
    request: DocxProposalStateRequest,
    main: (request: DocxProposalStateRequest) => Promise<DocxProposalResult>
  ): Promise<DocxProposalResult>;
  withdraw(
    request: DocxProposalWithdrawRequest,
    main: (request: DocxProposalWithdrawRequest) => Promise<DocxProposalResult>
  ): Promise<DocxProposalResult>;
  getProposals(main: () => Promise<DocxProposalSnapshot>): Promise<DocxProposalSnapshot>;
  readParagraphs(
    request: DocxReadParagraphsRequest,
    main: (request: DocxReadParagraphsRequest) => Promise<DocxReadParagraphsResult>
  ): Promise<DocxReadParagraphsResult>;
  paragraphIdentities(
    main: () => Promise<DocxParagraphIdentitySnapshot>
  ): Promise<DocxParagraphIdentitySnapshot>;
  resolveParagraphAnchors(
    anchors: readonly DocxParagraphAnchor[],
    main: (anchors: readonly DocxParagraphAnchor[]) => Promise<{
      version: string;
      results: DocxParagraphAnchorResult[];
    }>
  ): Promise<{ version: string; results: DocxParagraphAnchorResult[] }>;
  /** Resolves against the document the host sees now. */
  navigationTarget(
    story: string,
    paraId: string,
    main: () => ReturnType<typeof resolveNavigationTarget>
  ): Promise<{ version: string; target: ReturnType<typeof resolveNavigationTarget> }>;
  /** Geometry, initialization or hand-over changed. */
  subscribe(listener: () => void): () => void;
}

type Handover = { state: Uint8Array; complete(): void };
type RegisteredAuthority = WorkerProposalAuthority & {
  beginHandover(): Promise<Handover>;
  draining(): boolean;
  fail(error: unknown): void;
  failure(): unknown;
  handedOverRequest<T extends { expectVersion: string }>(request: T): T;
};
const authorities = new WeakMap<YrsSession, RegisteredAuthority>();

export function registerWorkerProposalAuthority(
  session: YrsSession,
  worker: Pick<WorkerOpenedDocument, 'proposal' | 'documentRead' | 'handOver'>,
  hooks: {
    relayout(): void;
    current(): boolean;
    /** Settles once the worker has laid the document out; proposal requests wait for it. */
    laidOut(): Promise<void>;
    adopted(version: string): void;
    handedOver(version: string): void;
    /** A worker proposal changed document content. */
    contentChanged(): void;
  }
): WorkerProposalAuthority {
  let tail: Promise<unknown> = Promise.resolve();
  let initializing: Promise<void> | null = null;
  let queued = 0;
  let snapshotPosted = false;
  let stopWaiting!: () => void;
  let failWaiting!: (error: unknown) => void;
  const stopped = new Promise<void>((resolve, reject) => {
    stopWaiting = resolve;
    failWaiting = reject;
  });
  void stopped.catch(() => {});
  let failure: { error: unknown } | null = null;
  const pendingCalls = new Set<(error: unknown) => void>();
  let initialized = false;
  let mirror: ResidentProposalReply['mirror'] | null = null;
  let geometry: ProposalGeometryMirror | null = null;
  let holdsState = false;
  let mutating = 0;
  let handingOver = false;
  let handover: Promise<Handover> | null = null;
  let transfer: (() => Promise<Handover>) | null = null;
  let versionRewrite: { worker: string; main: string } | null = null;
  const listeners = new Set<() => void>();
  const notify = () => { for (const listener of listeners) listener(); };
  const assertCurrent = () => {
    if (failure) throw failure.error;
    if (!hooks.current()) throw new Error('The document changed while applying proposals');
  };
  const interruptible = <T>(pending: Promise<T>): Promise<T> => {
    let rejectFailed!: (error: unknown) => void;
    const failed = new Promise<never>((_, reject) => { rejectFailed = reject; });
    pendingCalls.add(rejectFailed);
    if (failure) rejectFailed(failure.error);
    return Promise.race([pending, failed]).finally(() => { pendingCalls.delete(rejectFailed); });
  };
  const enqueue = <T>(call: () => Promise<T>): Promise<T> => {
    queued += 1;
    const result = tail.then(() => {
      assertCurrent();
      return interruptible(call());
    });
    const finished = () => { queued -= 1; };
    tail = result.then(finished, finished);
    return result;
  };
  const store = (reply: ResidentProposalReply): void => {
    // Listeners the mirror notifies read the geometry that goes with it.
    geometry = reply.geometry;
    mirror = reply.mirror;
    session.mirrorWorkerDocument(mirror);
    notify();
  };
  // Calls answer in order. One made before the hand-over began runs in the worker, ahead of the
  // hand-over; one made after waits in turn for the replica.
  const route = <T>(call: () => Promise<T>, main: () => T | Promise<T>): Promise<T> => {
    const ready = authority.initialize();
    void ready.catch(() => {});
    const viaWorker = !handingOver;
    return enqueue(async () => {
      if (viaWorker) {
        await ready;
        assertCurrent();
        if (!initialized && !handingOver) await initializeNow();
        if (initialized) return call();
      }
      // A hand-over queued behind this call runs now: nothing ahead of it is left for the worker.
      if (transfer) void transfer().catch(() => {});
      await awaitWorkerOpenReplica(session);
      assertCurrent();
      return main();
    });
  };
  const mutate = (
    op: Parameters<WorkerOpenedDocument['proposal']>[0],
    main: () => Promise<DocxProposalResult>
  ): Promise<DocxProposalResult> => route(async () => {
    const previous = mirror!;
    const previousPreview = JSON.stringify(proposalRevisionPreview(session.getProposals()));
    const pending = op.kind === 'propose' || op.kind === 'withdraw';
    if (pending) session.mirrorWorkerDocument({ ...previous, version: previous.version + '~' });
    let reply: ResidentProposalReply;
    mutating += 1;
    try {
      reply = await worker.proposal(op);
    } catch (error) {
      if (hooks.current() && pending && workerOpenReplicaPending(session)) {
        session.mirrorWorkerDocument(previous);
      }
      throw error;
    } finally {
      mutating -= 1;
    }
    assertCurrent();
    if (
      reply.changedStories.length > 0 ||
      JSON.stringify(previous.proposals) !== JSON.stringify(reply.mirror.proposals) ||
      (op.kind === 'setStates' && reply.result?.ok)
    ) holdsState = true;
    store(reply);
    if (
      reply.changedStories.length > 0 ||
      JSON.stringify(proposalRevisionPreview(session.getProposals())) !== previousPreview
    ) hooks.relayout();
    if (reply.changedStories.length > 0) hooks.contentChanged();
    if (!reply.result) throw new Error('The resident worker did not return a proposal result');
    return reply.result;
  }, main);
  const initializeNow = async (): Promise<void> => {
    await Promise.race([hooks.laidOut(), stopped]);
    assertCurrent();
    if (handingOver || initialized) return;
    snapshotPosted = true;
    const reply = await worker.proposal({ kind: 'snapshot' });
    assertCurrent();
    const previousVersion = mirror?.version;
    initialized = true;
    store(reply);
    hooks.adopted(reply.mirror.version);
    if (previousVersion !== undefined && previousVersion !== reply.mirror.version) hooks.relayout();
  };
  const authority: RegisteredAuthority = {
    get initialized() { return initialized; },
    restart() {
      if (!initialized || holdsState || failure || handingOver || !hooks.current()) return;
      initialized = false;
      initializing = null;
      snapshotPosted = false;
      hooks.relayout();
      notify();
    },
    initialize() {
      if (failure) return Promise.reject(failure.error);
      if (initializing) return initializing;
      if (handingOver) return Promise.resolve(awaitWorkerOpenReplica(session));
      initializing = enqueue(initializeNow);
      return initializing;
    },
    geometry: () => geometry,
    holdsWorkerState: () => holdsState || mutating > 0,
    holdsCommittedWorkerState: () => holdsState,
    failure: () => failure?.error,
    draining: () => handingOver && queued > 0,
    fail: (error) => {
      if (failure) return;
      failure = { error };
      failWaiting(error);
      for (const reject of pendingCalls) reject(error);
    },
    propose: (request, main) => mutate({ kind: 'propose', request }, () => main(request)),
    setStates: (request, main) => mutate({ kind: 'setStates', request }, () => main(request)),
    withdraw: (request, main) => mutate({ kind: 'withdraw', request }, () => main(request)),
    handedOverRequest: (request) =>
      versionRewrite &&
      request.expectVersion === versionRewrite.worker &&
      session.version() === versionRewrite.main
        ? { ...request, expectVersion: versionRewrite.main }
        : request,
    getProposals: (main) => route(async () => session.getProposals(), main),
    readParagraphs: (request, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'readParagraphs', request });
      assertCurrent();
      return read.value;
    }, () => main(request)),
    paragraphIdentities: (main) => route(async () => {
      const read = await worker.documentRead({ kind: 'paragraphIdentities' });
      assertCurrent();
      return read.value;
    }, main),
    resolveParagraphAnchors: (anchors, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'resolveParagraphAnchors', anchors: [...anchors] });
      assertCurrent();
      return { version: read.version, results: read.value.results };
    }, () => main(anchors)),
    navigationTarget: (story, paraId, main) => route(async () => {
      const local = resolveMirroredNavigationTarget(geometry, session.getProposals(), story, paraId);
      if (local !== null) return { version: geometry!.version, target: local };
      const read = await worker.documentRead({ kind: 'navigationTarget', story, paraId });
      assertCurrent();
      return { version: read.version, target: read.value };
    }, () => ({ version: session.version(), target: main() })),
    subscribe(listener) {
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    beginHandover() {
      if (failure) return Promise.reject(failure.error);
      if (handover) return handover;
      handingOver = true;
      stopWaiting();
      let transferring: Promise<Handover> | null = null;
      const transferOnce = async (): Promise<Handover> => {
        assertCurrent();
        const handedOver = await worker.handOver();
        assertCurrent();
        let completed = false;
        return {
          state: handedOver.state,
          complete() {
            if (completed) return;
            assertCurrent();
            completed = true;
            hooks.handedOver(handedOver.version);
            geometry = null;
            if (mirror) {
              session.mirrorWorkerDocument({ version: handedOver.version, proposals: handedOver.proposals });
              session.mirrorWorkerDocument(null);
              versionRewrite = { worker: handedOver.version, main: session.version() };
            }
            holdsState = false;
            notify();
          },
        };
      };
      let started!: () => void;
      const startedEarly = new Promise<void>((resolve) => { started = resolve; });
      const start = (): Promise<Handover> => {
        started();
        return (transferring ??= transferOnce());
      };
      transfer = start;
      handover = snapshotPosted
        ? Promise.race([enqueue(start), interruptible(startedEarly.then(() => transferring!))])
        : interruptible(start());
      return handover;
    },
  };
  authorities.set(session, authority);
  return authority;
}

export function workerProposalAuthority(session: YrsSession): WorkerProposalAuthority | null {
  const authority = authorities.get(session);
  return authority && (authority.draining() || (workerOpenReplicaPending(session) &&
    (!workerOpenReplicaStarted(session) || authority.holdsWorkerState())))
    ? authority
    : null;
}

export function registeredWorkerProposalAuthority(session: YrsSession): WorkerProposalAuthority | null {
  return authorities.get(session) ?? null;
}

export function workerProposalFailure(session: YrsSession): unknown {
  return authorities.get(session)?.failure();
}

/**
 * `request` with the worker version it was read at replaced by the main session's, while main
 * still holds exactly the state it took over from the worker.
 */
export function handedOverRequest<T extends { expectVersion: string }>(
  session: YrsSession,
  request: T
): T {
  return authorities.get(session)?.handedOverRequest(request) ?? request;
}

export function beginWorkerProposalHandover(session: YrsSession): Promise<Handover> | null {
  return authorities.get(session)?.beginHandover() ?? null;
}

export function failWorkerProposalAuthority(session: YrsSession, error: unknown): void {
  authorities.get(session)?.fail(error);
}

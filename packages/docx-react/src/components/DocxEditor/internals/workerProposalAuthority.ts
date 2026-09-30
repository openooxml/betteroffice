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
import {
  awaitWorkerOpenReplica,
  workerOpenReplicaPending,
  workerOpenReplicaStarted,
} from './workerOpenReplica';

export interface WorkerProposalAuthority {
  /** The session mirrors the worker registry and version. */
  readonly initialized: boolean;
  /** Initializes once; rejects when the worker cannot answer. */
  initialize(): Promise<void>;
  /** Mirrored geometry until hand-over. */
  geometry(): ProposalGeometryMirror | null;
  /** Reseeding would lose worker changes. */
  holdsWorkerState(): boolean;
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
  let initialized = false;
  let mirror: ResidentProposalReply['mirror'] | null = null;
  let geometry: ProposalGeometryMirror | null = null;
  let holdsState = false;
  let handingOver = false;
  let handover: Promise<Handover> | null = null;
  let versionRewrite: { worker: string; main: string } | null = null;
  const listeners = new Set<() => void>();
  const notify = () => { for (const listener of listeners) listener(); };
  const assertCurrent = () => {
    if (!hooks.current()) throw new Error('The document changed while applying proposals');
  };
  const enqueue = <T>(call: () => Promise<T>): Promise<T> => {
    const result = tail.then(call);
    tail = result.then(() => {}, () => {});
    return result;
  };
  const store = (reply: ResidentProposalReply): void => {
    mirror = reply.mirror;
    session.mirrorWorkerDocument(mirror);
    geometry = reply.geometry;
    notify();
  };
  // Calls answer in order, whichever side takes them.
  const route = <T>(call: () => Promise<T>, main: () => T | Promise<T>): Promise<T> => {
    const ready = authority.initialize();
    void ready.catch(() => {});
    return enqueue(async () => {
      if (!handingOver) {
        await ready;
        assertCurrent();
        return call();
      }
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
    const pending = op.kind === 'propose' || op.kind === 'withdraw';
    if (pending) session.mirrorWorkerDocument({ ...previous, version: previous.version + '~' });
    let reply: ResidentProposalReply;
    try {
      reply = await worker.proposal(op);
    } catch (error) {
      if (hooks.current() && pending) session.mirrorWorkerDocument(previous);
      throw error;
    }
    assertCurrent();
    if (
      reply.changedStories.length > 0 ||
      JSON.stringify(previous.proposals) !== JSON.stringify(reply.mirror.proposals) ||
      (op.kind === 'setStates' && reply.result?.ok)
    ) holdsState = true;
    store(reply);
    if (pending && reply.changedStories.length > 0) hooks.relayout();
    if (reply.changedStories.length > 0) hooks.contentChanged();
    if (!reply.result) throw new Error('The resident worker did not return a proposal result');
    return reply.result;
  }, main);
  const authority: RegisteredAuthority = {
    get initialized() { return initialized; },
    initialize() {
      if (initializing) return initializing;
      if (handingOver) return Promise.resolve(awaitWorkerOpenReplica(session));
      initializing = enqueue(async () => {
        await hooks.laidOut();
        assertCurrent();
        const reply = await worker.proposal({ kind: 'snapshot' });
        assertCurrent();
        initialized = true;
        store(reply);
        hooks.adopted(reply.mirror.version);
      });
      return initializing;
    },
    geometry: () => geometry,
    holdsWorkerState: () => holdsState,
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
      const read = await worker.documentRead({ kind: 'navigationTarget', story, paraId });
      assertCurrent();
      return { version: read.version, target: read.value };
    }, () => ({ version: session.version(), target: main() })),
    subscribe(listener) {
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    beginHandover() {
      if (handover) return handover;
      handingOver = true;
      handover = enqueue(async () => {
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
            if (initialized) {
              session.mirrorWorkerDocument({ version: handedOver.version, proposals: handedOver.proposals });
              session.mirrorWorkerDocument(null);
              versionRewrite = { worker: handedOver.version, main: session.version() };
            }
            geometry = null;
            notify();
          },
        };
      });
      return handover;
    },
  };
  authorities.set(session, authority);
  return authority;
}

export function workerProposalAuthority(session: YrsSession): WorkerProposalAuthority | null {
  const authority = authorities.get(session);
  return authority && workerOpenReplicaPending(session) &&
    (!workerOpenReplicaStarted(session) || authority.holdsWorkerState())
    ? authority
    : null;
}

export function registeredWorkerProposalAuthority(session: YrsSession): WorkerProposalAuthority | null {
  return authorities.get(session) ?? null;
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

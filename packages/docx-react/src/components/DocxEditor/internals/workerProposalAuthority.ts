import type {
  DocxContentControlQuery,
  DocxContentControlsOptions,
  DocxContentControlsResult,
  DocxExportResult,
  DocxLayoutMap,
  DocxPageExportOptions,
  DocxPagedStructuredContent,
  DocxFindTextRequest,
  DocxFindTextResult,
  DocxFindParagraphsOptions,
  DocxParagraphMatch,
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
  ResidentEngineWorkerClient,
  YrsLoc,
  YrsSession,
  YrsStickyPosition,
  resolveNavigationTarget,
} from '@betteroffice/docx/yrs';
import type { WorkerOpenedDocument } from '../hooks/useDisplayList';
import { proposalRevisionPreview, resolveMirroredNavigationTarget } from '@betteroffice/docx/yrs';
import {
  awaitWorkerOpenReplica,
  adoptWorkerOpenHandoverVersion,
  workerOpenReplicaPending,
  workerOpenReplicaReady,
  workerOpenReplicaStarted,
  workerOpenDocumentHeld,
  workerOpenRequest,
} from './workerOpenReplica';

type SearchRead = Awaited<ReturnType<typeof ResidentEngineWorkerClient.prototype.documentRead<'searchText'>>>;
type StickyAnchorsRead = Awaited<ReturnType<typeof ResidentEngineWorkerClient.prototype.documentRead<'stickyAnchors'>>>;

export interface WorkerProposalAuthority {
  /** The authority has a worker registry snapshot. */
  readonly initialized: boolean;
  snapshot(): DocxProposalSnapshot | null;
  revisionPreview(): ReturnType<typeof proposalRevisionPreview>;
  previewVersion(): number;
  workerCoversPeer(peerVersion: string): boolean;
  retirementReason(): 'source-fallback' | null;
  retire(reason: 'source-fallback'): boolean;
  catchUp(complete?: () => void): Promise<void>;
  hydratePeer(load: () => Promise<void>, complete: () => void): Promise<void>;
  restart(): void;
  save<T>(task: () => Promise<T>): Promise<T>;
  residentOperation<T>(task: () => Promise<T>): Promise<T>;
  /** Initializes once; rejects when the worker cannot answer. */
  initialize(): Promise<void>;
  /** Geometry of the worker registry. */
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
  removeComment(id: string, main: () => void | Promise<void>): Promise<void>;
  getProposals(main: () => Promise<DocxProposalSnapshot>): Promise<DocxProposalSnapshot>;
  readParagraphs(
    request: DocxReadParagraphsRequest,
    main: (request: DocxReadParagraphsRequest) => Promise<DocxReadParagraphsResult>
  ): Promise<DocxReadParagraphsResult>;
  paragraphIdentities(
    main: () => Promise<DocxParagraphIdentitySnapshot>
  ): Promise<DocxParagraphIdentitySnapshot>;
  searchText(
    query: string,
    caseSensitive: boolean,
    carry: YrsStickyPosition | null,
    main: () => SearchRead['value']
  ): Promise<SearchRead>;
  stickyAnchors(
    locs: YrsLoc[],
    version: string,
    main: () => Array<YrsStickyPosition | null>
  ): Promise<StickyAnchorsRead>;
  resolveParagraphAnchors(
    anchors: readonly DocxParagraphAnchor[],
    main: (anchors: readonly DocxParagraphAnchor[]) => Promise<{
      version: string;
      results: DocxParagraphAnchorResult[];
    }>
  ): Promise<{ version: string; results: DocxParagraphAnchorResult[] }>;
  findParagraphs(
    query: string,
    options: DocxFindParagraphsOptions | undefined,
    main: () => Promise<DocxParagraphMatch[]>
  ): Promise<DocxParagraphMatch[]>;
  findText(
    request: DocxFindTextRequest,
    main: () => Promise<DocxFindTextResult>
  ): Promise<DocxFindTextResult>;
  listContentControls(
    options: DocxContentControlsOptions | undefined,
    main: () => Promise<DocxContentControlsResult>
  ): Promise<DocxContentControlsResult>;
  findContentControls(
    query: DocxContentControlQuery,
    options: DocxContentControlsOptions | undefined,
    main: () => Promise<DocxContentControlsResult>
  ): Promise<DocxContentControlsResult>;
  /**
   * Reads the layout request after the calls ahead of it, then the worker's export for it; null
   * when there is no request yet.
   */
  exportStructuredWithPages(
    options: DocxPageExportOptions,
    currentRequest: () => Promise<string | null>,
    main: () => Promise<DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>>
  ): Promise<DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>> | null>;
  /** Resolves against the document the host sees now. */
  navigationTarget(
    story: string,
    paraId: string,
    main: () => ReturnType<typeof resolveNavigationTarget>
  ): Promise<{ version: string; target: ReturnType<typeof resolveNavigationTarget> }>;
  /** Geometry or initialization changed. */
  subscribe(listener: () => void): () => void;
}

type PeerSnapshot = Pick<Awaited<ReturnType<WorkerOpenedDocument['handOver']>>,
  'state' | 'version' | 'metadata' | 'metadataReason'>;
type RegisteredAuthority = WorkerProposalAuthority & {
  readsInWorker(): boolean;
  peerSnapshot(): Promise<PeerSnapshot>;
  fail(error: unknown): void;
  failure(): unknown;
  handedOverRequest<T extends { expectVersion: string }>(request: T): T;
};
class StaleProposalVersionError extends Error {
  constructor(readonly version: string) { super('the document changed since the expected version was read'); }
}
const authorities = new WeakMap<YrsSession, RegisteredAuthority>();

export function registerWorkerProposalAuthority(
  session: YrsSession,
  worker: Pick<WorkerOpenedDocument, 'proposal' | 'documentRead' | 'handOver'> &
    Partial<Pick<WorkerOpenedDocument, 'syncUpdate' | 'integrateProposalUpdate' | 'stateRevision'>>,
  hooks: {
    editorPeer?: boolean;
    passiveEditor?: boolean;
    relayout(): void;
    current(): boolean;
    /** Settles once the worker has laid the document out; proposal requests wait for it. */
    laidOut(): Promise<void>;
    adopted(version: string): void;
    handedOver?(version: string): void;
    /** A worker proposal changed document content. */
    contentChanged(): void;
    projectionChanged?(stories: readonly string[]): void;
    peerUpdated?(stories: readonly string[]): void;
  }
): WorkerProposalAuthority {
  let tail: Promise<unknown> = Promise.resolve();
  let initializing: Promise<void> | null = null;
  let failWaiting!: (error: unknown) => void;
  const stopped = new Promise<void>((_resolve, reject) => {
    failWaiting = reject;
  });
  void stopped.catch(() => {});
  let failure: { error: unknown } | null = null;
  const pendingCalls = new Set<(error: unknown) => void>();
  let initialized = false;
  let layoutRevision = 0;
  let mirror: ResidentProposalReply['mirror'] | null = null;
  let geometry: ProposalGeometryMirror | null = null;
  let holdsState = false;
  let mutating = 0;
  let peerHydrated = false;
  let retirementReason: 'source-fallback' | null = null;
  let releaseRetirement!: () => void;
  const retired = new Promise<void>((resolve) => { releaseRetirement = resolve; });
  let correspondence: {
    worker: string; peer: string; revision?: ReturnType<NonNullable<WorkerOpenedDocument['stateRevision']>>;
  } | null = null;
  const editorPeer = (): boolean => hooks.editorPeer === true && !workerOpenDocumentHeld(session);
  const peerReady = (): boolean => peerHydrated || (editorPeer() && workerOpenReplicaReady(session));
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
  const enqueue = <T>(call: () => Promise<T>, allowFailedPeer = false): Promise<T> => {
    const result = tail.then(() => {
      if (allowFailedPeer && failure && editorPeer() && peerReady() && !holdsState && mutating === 0 && hooks.current()) {
        return call();
      }
      assertCurrent();
      return interruptible(call());
    });
    tail = result.then(() => {}, () => {});
    return result;
  };
  const store = (reply: ResidentProposalReply): void => {
    // Listeners the mirror notifies read the geometry that goes with it.
    geometry = peerReady() ? { ...reply.geometry, version: session.version() } : reply.geometry;
    if (correspondence?.worker !== reply.mirror.version) correspondence = null;
    mirror = reply.mirror;
    if (!peerReady()) session.mirrorWorkerDocument(mirror, !editorPeer());
    notify();
  };
  const route = <T>(call: () => Promise<T>, main: () => T | Promise<T>): Promise<T> => {
    if (!workerProposalAuthority(session)) {
      return Promise.resolve(awaitWorkerOpenReplica(session)).then(() => enqueue(async () => {
        assertCurrent();
        return main();
      }));
    }
    const ready = authority.initialize();
    void ready.catch(() => {});
    let execution: Promise<T> | null = null;
    const execute = (workerAdmitted: boolean): Promise<T> => execution ??= enqueue(async () => {
      if (!workerAdmitted || retirementReason) return main();
      assertCurrent();
      return call();
    });
    if (initialized) return execute(true);
    const workerRead = ready.then(() => execute(true));
    const loaded = workerOpenDocumentHeld(session) ? undefined : awaitWorkerOpenReplica(session);
    return loaded ? Promise.race([workerRead, loaded.then(() => execute(false))]) : workerRead;
  };
  const round = <T>(call: () => Promise<T>, main: () => T | Promise<T>): Promise<T> => {
    const ready = authority.initialize();
    void ready.catch(() => {});
    const admittedRevision = layoutRevision;
    const execute = (revision: number): Promise<T> => enqueue<{ value: T } | { readmit: true }>(async () => {
      if (retirementReason) return { value: await main() };
      await ready;
      assertCurrent();
      if (retirementReason) return { value: await main() };
      if (!initialized && revision !== layoutRevision) return { readmit: true };
      if (!initialized) await initializeNow();
      if (retirementReason) return { value: await main() };
      return { value: await call() };
    }).then((result) => {
      if ('value' in result) return result.value;
      const nextRevision = layoutRevision;
      return Promise.race([hooks.laidOut(), stopped, retired]).then(() => execute(nextRevision));
    });
    return initialized ? execute(admittedRevision) : ready.then(() => execute(admittedRevision));
  };
  const workerSnapshot = (): DocxProposalSnapshot | null => retirementReason ? session.getProposals() : mirror ? {
    version: peerReady() ? session.version() : mirror.version,
    previewVersion: mirror.proposals.previewVersion,
    proposals: mirror.proposals.entries.map(({ record }) => ({
      ...record, paragraph: { ...record.paragraph }, revisionIds: [...record.revisionIds],
    })),
  } : null;
  const preview = (): ReturnType<typeof proposalRevisionPreview> => {
    if (retirementReason) return proposalRevisionPreview(session.getProposals());
    const snapshot = workerSnapshot();
    const workerPreview = snapshot ? proposalRevisionPreview(snapshot) : undefined;
    if (!editorPeer() || !peerReady()) return workerPreview;
    const combined = { ...workerPreview, ...proposalRevisionPreview(session.getProposals()) };
    return Object.keys(combined).length === 0 ? undefined : Object.fromEntries(
      Object.entries(combined).sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0)
    );
  };
  const integrate = (diff: Uint8Array, stories: readonly string[]): void => {
    if (!worker.integrateProposalUpdate) throw new Error('The worker proposal peer integration is unavailable');
    const integrated = worker.integrateProposalUpdate(diff, stories);
    const changed: readonly string[] = Array.isArray(integrated) ? integrated : stories;
    hooks.projectionChanged?.(changed);
    hooks.peerUpdated?.(changed);
  };
  const recordCorrespondence = (workerVersion: string, unchanged: boolean, vector: Uint8Array): void => {
    const peerVector = session.encodeStateVector();
    correspondence = unchanged && peerVector.length === vector.length &&
      peerVector.every((byte, index) => byte === vector[index])
      ? { worker: workerVersion, peer: session.version(), revision: worker.stateRevision?.() }
      : null;
  };
  const sendMutation = async (
    op: Parameters<WorkerOpenedDocument['proposal']>[0],
    prepare?: () => Parameters<WorkerOpenedDocument['proposal']>[0]
  ): Promise<ResidentProposalReply> => {
    const previous = mirror!;
    const previousPreview = JSON.stringify(proposalRevisionPreview(workerSnapshot()!));
    const pending = op.kind === 'propose' || op.kind === 'withdraw' || op.kind === 'removeComment';
    if (pending && !peerReady()) session.mirrorWorkerDocument(
      { ...previous, version: previous.version + '~' }, !editorPeer()
    );
    let postedVersion = editorPeer() ? session.version() : undefined;
    let reply: ResidentProposalReply;
    mutating += 1;
    try {
      reply = await (prepare ? worker.proposal(op, () => {
        const operation = prepare();
        if (editorPeer()) postedVersion = session.version();
        return operation;
      }) : worker.proposal(op));
    } catch (error) {
      if (hooks.current() && pending && workerOpenReplicaPending(session)) {
        session.mirrorWorkerDocument(previous, !editorPeer());
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
    if (peerReady() && reply.peerDiff) {
      const unchanged = session.version() === postedVersion;
      integrate(reply.peerDiff, [...new Set([...reply.changedStories, ...(reply.projectionStories ?? [])])]);
      recordCorrespondence(reply.mirror.version, unchanged, reply.stateVector);
    } else {
      if (peerReady()) correspondence = null;
      hooks.projectionChanged?.(reply.projectionStories ?? []);
    }
    store(reply);
    if (
      reply.changedStories.length > 0 ||
      JSON.stringify(proposalRevisionPreview(workerSnapshot()!)) !== previousPreview
    ) hooks.relayout();
    if (reply.changedStories.length > 0) hooks.contentChanged();
    return reply;
  };
  const mutate = (
    op: Extract<Parameters<WorkerOpenedDocument['proposal']>[0], { request: unknown }>,
    main: () => Promise<DocxProposalResult>
  ): Promise<DocxProposalResult> => round<DocxProposalResult>(async () => {
    const prepare = () => {
      if (!editorPeer()) return op;
      const version = peerReady() ? session.version() : mirror!.version;
      const token = op.request.expectVersion;
      if (token !== version && !(correspondence?.peer === version && correspondence.worker === token)) {
        throw new StaleProposalVersionError(version);
      }
      return { ...op, peerStateVector: session.encodeStateVector() };
    };
    let reply: ResidentProposalReply;
    try {
      reply = await sendMutation(prepare(), editorPeer() ? prepare : undefined);
    } catch (error) {
      if (!(error instanceof StaleProposalVersionError)) throw error;
      return { ok: false, version: error.version, failure: { code: 'stale-version', message: error.message } };
    }
    if (!reply.result) throw new Error('The resident worker did not return a proposal result');
    if (!peerReady()) return reply.result;
    return reply.result.ok
      ? { ...reply.result, snapshot: { ...reply.result.snapshot, version: session.version() } }
      : { ...reply.result, version: session.version() };
  }, main);
  const initializeNow = async (): Promise<void> => {
    if (retirementReason) return;
    assertCurrent();
    if (initialized || retirementReason) return;
    const reply = await Promise.race([worker.proposal({ kind: 'snapshot' }), retired.then(() => null)]).catch((error: unknown) => {
      if (retirementReason) return null;
      throw error;
    });
    assertCurrent();
    if (!reply || retirementReason) return;
    const previousVersion = mirror?.version;
    initialized = true;
    store(reply);
    hooks.adopted(reply.mirror.version);
    if (previousVersion !== undefined && previousVersion !== reply.mirror.version) hooks.relayout();
  };
  const catchUpNow = async (complete?: () => void): Promise<void> => {
    if (retirementReason) { complete?.(); return; }
    if (!worker.syncUpdate) throw new Error('The worker proposal catch-up is unavailable');
    const version = session.version();
    const reply = await worker.syncUpdate(new Uint8Array(), session.encodeStateVector());
    assertCurrent();
    const unchanged = session.version() === version;
    if (reply.repair) integrate(reply.repair, []);
    peerHydrated = true;
    adoptWorkerOpenHandoverVersion(session, reply.version);
    session.mirrorWorkerDocument(null, false);
    recordCorrespondence(reply.version, unchanged, reply.stateVector);
    if (initialized) {
      const snapshot = await worker.proposal({ kind: 'snapshot' });
      assertCurrent();
      initializing = Promise.resolve();
      store(snapshot);
      correspondence = snapshot.mirror.version === reply.version ? correspondence : null;
    } else {
      initializing = null;
      notify();
    }
    if (correspondence && mirror?.proposals.entries.length === 0 && mirror.proposals.previewVersion === 0) {
      holdsState = false;
    }
    complete?.();
  };
  const authority: RegisteredAuthority = {
    readsInWorker: () => !hooks.passiveEditor || workerOpenDocumentHeld(session) || holdsState || mutating > 0,
    get initialized() { return initialized; },
    snapshot: workerSnapshot,
    revisionPreview: preview,
    previewVersion: () => (mirror?.proposals.previewVersion ?? 0) +
      (retirementReason || (editorPeer() && peerReady()) ? session.getProposals().previewVersion : 0),
    workerCoversPeer: (version) => {
      if (retirementReason || mutating > 0 || correspondence?.peer !== version || session.version() !== version) return false;
      const revision = worker.stateRevision?.();
      return revision?.owner === correspondence.revision?.owner && revision?.sequence === correspondence.revision?.sequence;
    },
    retirementReason: () => retirementReason,
    retire(reason) {
      if (retirementReason) return false;
      if (holdsState || mutating > 0) throw new Error('The resident worker holds proposals the main thread cannot rebuild');
      retirementReason = reason;
      failure = null;
      correspondence = null;
      mirror = null;
      geometry = null;
      initialized = false;
      releaseRetirement();
      notify();
      return true;
    },
    catchUp: (complete) => enqueue(() => catchUpNow(complete)),
    hydratePeer: (load, complete) => enqueue(async () => {
      await load();
      assertCurrent();
      await catchUpNow(complete);
    }),
    residentOperation: (task) => enqueue(async () => {
      const result = await task();
      assertCurrent();
      return result;
    }),
    restart() {
      if (retirementReason || !initialized || holdsState || failure || !hooks.current()) return;
      initialized = false;
      layoutRevision += 1;
      initializing = null;
      hooks.relayout();
      notify();
    },
    save: (task) => enqueue(async () => {
      const previous = mirror?.version;
      const result = await task();
      assertCurrent();
      if (initialized && !retirementReason) {
        const reply = await worker.proposal({ kind: 'snapshot' });
        assertCurrent();
        store(reply);
        if (reply.mirror.version !== previous) hooks.relayout();
      }
      return result;
    }, true),
    initialize() {
      if (failure) return Promise.reject(failure.error);
      if (retirementReason) return Promise.resolve();
      if (initializing) return initializing;
      initializing = Promise.race([hooks.laidOut(), stopped, retired]).then(() => enqueue(initializeNow));
      return initializing;
    },
    geometry: () => geometry,
    holdsWorkerState: () => holdsState || mutating > 0,
    holdsCommittedWorkerState: () => holdsState,
    failure: () => failure?.error,
    fail: (error) => {
      if (failure) return;
      failure = { error };
      failWaiting(error);
      for (const reject of pendingCalls) reject(error);
    },
    propose: (request, main) => mutate({ kind: 'propose', request }, () => main(request)),
    setStates: (request, main) => mutate({ kind: 'setStates', request }, () => main(request)),
    withdraw: (request, main) => mutate({ kind: 'withdraw', request }, () => main(request)),
    removeComment: (id, main) => route(async () => {
      await sendMutation({ kind: 'removeComment', id });
    }, main),
    handedOverRequest: (request) =>
      !retirementReason && correspondence && request.expectVersion === correspondence.worker &&
      session.version() === correspondence.peer
        ? { ...request, expectVersion: correspondence.peer }
        : request,
    getProposals: (main) => round(async () => {
      if (peerReady() && editorPeer()) {
        const reply = await worker.proposal({ kind: 'snapshot' });
        assertCurrent();
        store(reply);
      }
      return workerSnapshot()!;
    }, main),
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
    searchText: (query, caseSensitive, carry, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'searchText', query, caseSensitive, carry });
      assertCurrent();
      return read;
    }, () => ({ version: session.version(), value: main() })),
    stickyAnchors: (locs, version, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'stickyAnchors', locs, version });
      assertCurrent();
      return read;
    }, () => ({ version: session.version(), value: main() })),
    resolveParagraphAnchors: (anchors, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'resolveParagraphAnchors', anchors: [...anchors] });
      assertCurrent();
      return { version: read.version, results: read.value.results };
    }, () => main(anchors)),
    findParagraphs: (query, options, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'findParagraphs', query, ...options });
      assertCurrent();
      return read.value;
    }, main),
    findText: (request, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'findText', request });
      assertCurrent();
      return read.value;
    }, main),
    listContentControls: (options, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'listContentControls', options: options ?? {} });
      assertCurrent();
      return read.value;
    }, main),
    findContentControls: (query, options, main) => route(async () => {
      const read = await worker.documentRead({ kind: 'findContentControls', query, options: options ?? {} });
      assertCurrent();
      return read.value;
    }, main),
    exportStructuredWithPages: (options, currentRequest, main) => route<
      DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>> | null
    >(async () => {
      const request = await currentRequest();
      assertCurrent();
      if (request === null) return null;
      const read = await worker.documentRead({ kind: 'exportStructuredWithPages', options, currentRequest: request });
      assertCurrent();
      return JSON.parse(read.value) as DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>;
    }, main),
    navigationTarget: (story, paraId, main) => route(async () => {
      const local = resolveMirroredNavigationTarget(geometry, workerSnapshot()!, story, paraId);
      if (local !== null) return { version: geometry!.version, target: local };
      const read = await worker.documentRead({ kind: 'navigationTarget', story, paraId });
      assertCurrent();
      return { version: read.version, target: read.value };
    }, () => ({ version: session.version(), target: main() })),
    subscribe(listener) {
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    peerSnapshot: () => enqueue(async () => {
      const snapshot = await worker.handOver();
      assertCurrent();
      return {
        state: snapshot.state,
        version: snapshot.version,
        ...(snapshot.metadata === undefined ? {} : { metadata: snapshot.metadata }),
        ...(snapshot.metadataReason === undefined ? {} : { metadataReason: snapshot.metadataReason }),
      };
    }),
  };
  authorities.set(session, authority);
  return authority;
}

export function workerProposalAuthority(session: YrsSession, includeHydrating = false): WorkerProposalAuthority | null {
  const authority = authorities.get(session);
  return authority && !authority.retirementReason() && authority.readsInWorker() && workerOpenReplicaPending(session) &&
    (!workerOpenReplicaStarted(session) || authority.holdsWorkerState() || (includeHydrating && authority.initialized))
    ? authority
    : null;
}

/**
 * Decision A: proposal rounds keep the resident worker authority after editor hydration.
 * Proposal ids are independent across the peer-local and worker registries.
 */
export function workerProposalRoundAuthority(session: YrsSession): WorkerProposalAuthority | null {
  return authorities.get(session) ?? null;
}

export const registeredWorkerProposalAuthority = workerProposalRoundAuthority;

export function workerProposalFailure(session: YrsSession): unknown {
  return authorities.get(session)?.failure();
}

/**
 * Maps a worker token only while the peer still holds its corresponding version.
 */
export function handedOverRequest<T extends { expectVersion: string }>(
  session: YrsSession,
  request: T
): T {
  const authority = authorities.get(session);
  return authority ? authority.handedOverRequest(request) : workerOpenRequest(session, request);
}

export function snapshotWorkerProposalPeer(session: YrsSession): Promise<PeerSnapshot> | null {
  return authorities.get(session)?.peerSnapshot() ?? null;
}

export function failWorkerProposalAuthority(session: YrsSession, error: unknown): void {
  authorities.get(session)?.fail(error);
}

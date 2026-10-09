import * as baseAuthority from './baseWorkerProposalAuthority';
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
  DocxProposalInput,
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
import { stateVectorAhead } from './stateVector';
import { proposalRevisionPreview, resolveMirroredNavigationTarget } from '@betteroffice/docx/yrs';
import {
  awaitWorkerOpenReplica,
  workerOpenReplicaPending,
  workerOpenReplicaReady,
  workerOpenReplicaStarted,
  workerOpenDocumentHeld,
  workerOpenRequest,
} from './workerOpenReplica';

type SearchRead = Awaited<ReturnType<typeof ResidentEngineWorkerClient.prototype.documentRead<'searchText'>>>;
type StickyAnchorsRead = Awaited<ReturnType<typeof ResidentEngineWorkerClient.prototype.documentRead<'stickyAnchors'>>>;

type RoundAdmission = { flush(): Promise<void>; validate(): void };

export interface WorkerProposalAuthority {
  /** The authority has a worker registry snapshot. */
  readonly initialized: boolean;
  snapshot(): DocxProposalSnapshot | null;
  revisionPreview(): ReturnType<typeof proposalRevisionPreview>;
  previewVersion(): number;
  workerCoversPeer(peerVersion: string): boolean;
  retirementReason(): 'source-fallback' | null;
  retire(reason: 'source-fallback'): boolean;
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
    main: (request: DocxProposalRequest) => Promise<DocxProposalResult>,
    admission?: RoundAdmission
  ): Promise<DocxProposalResult>;
  setStates(
    request: DocxProposalStateRequest,
    main: (request: DocxProposalStateRequest) => Promise<DocxProposalResult>,
    admission?: RoundAdmission
  ): Promise<DocxProposalResult>;
  withdraw(
    request: DocxProposalWithdrawRequest,
    main: (request: DocxProposalWithdrawRequest) => Promise<DocxProposalResult>,
    admission?: RoundAdmission
  ): Promise<DocxProposalResult>;
  removeComment(id: string, main: () => void | Promise<void>): Promise<void>;
  getProposals(main: () => Promise<DocxProposalSnapshot>, admission?: RoundAdmission): Promise<DocxProposalSnapshot>;
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
  hasEditorRounds(): boolean;
  readsInWorker(): boolean;
  peerSnapshot(): Promise<PeerSnapshot>;
  fail(error: unknown): void;
  failure(): unknown;
  handedOverRequest<T extends { expectVersion: string }>(request: T): T;
};
class StaleProposalVersionError extends Error {
  constructor(readonly version: string) { super('the document changed since the expected version was read'); }
}
function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value === null || typeof value !== 'object') return value;
  const record = value as Record<string, unknown>;
  return Object.fromEntries(Object.keys(record).sort().filter((key) => record[key] !== undefined)
    .map((key) => [key, canonical(record[key])]));
}

function retryKey(input: DocxProposalInput): string {
  const edit = input.op === 'replaceText'
    ? { op: input.op, search: input.search, replaceWith: input.replaceWith,
        occurrence: input.occurrence === undefined || input.occurrence === 1 ? 'first' : input.occurrence }
    : { op: input.op, at: input.at, text: input.text };
  return JSON.stringify(canonical({ paragraph: input.paragraph, ...edit }));
}

const authorities = new WeakMap<YrsSession, RegisteredAuthority>();

function sameVector(a: Uint8Array, b: Uint8Array): boolean {
  return (a.length === b.length && a.every((byte, index) => byte === b[index])) ||
    (a.length > 0 && b.length > 0 && !stateVectorAhead(a, b) && !stateVectorAhead(b, a));
}

export function registerWorkerProposalAuthority(
  session: YrsSession,
  worker: Pick<WorkerOpenedDocument, 'proposal' | 'documentRead' | 'handOver'> &
    Partial<Pick<WorkerOpenedDocument, 'syncUpdate' | 'integrateProposalUpdate' | 'stateRevision'>>,
  hooks: {
    editorPeer?: boolean;
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
  if (!hooks.editorPeer) {
    const authority = baseAuthority.registerWorkerProposalAuthority(session, worker, {
      ...hooks, handedOver: hooks.handedOver ?? (() => {}),
    });
    return extendBaseAuthority(session, authority);
  }
  let tail: Promise<unknown> = Promise.resolve();
  let initializing: Promise<void> | null = null;
  let initializationFailure: { error: unknown } | null = null;
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
  let admittedRounds = 0;
  const editorRounds = hooks.editorPeer === true;
  let retiredRegistry: ReturnType<YrsSession['createWorkerProposalRegistry']> | null = null;
  let retirementReason: 'source-fallback' | null = null;
  let releaseRetirement!: () => void;
  const retired = new Promise<void>((resolve) => { releaseRetirement = resolve; });
  let correspondence: {
    worker: string; peer: string; revision?: ReturnType<NonNullable<WorkerOpenedDocument['stateRevision']>>;
  } | null = null;
  const editorPeer = (): boolean => hooks.editorPeer === true && !workerOpenDocumentHeld(session);
  const peerReady = (): boolean => workerOpenReplicaReady(session);
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
    geometry = peerReady()
      ? sameVector(session.encodeStateVector(), reply.stateVector)
        ? { ...reply.geometry, version: session.version() }
        : null
      : reply.geometry;
    if (correspondence?.worker !== reply.mirror.version) correspondence = null;
    mirror = reply.mirror;
    if (!editorPeer() && !peerReady()) session.mirrorWorkerDocument(mirror);
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
    const workerRead = ready.then(async () => {
      if (!initialized) await awaitWorkerOpenReplica(session);
      return execute(initialized);
    });
    const loaded = workerOpenDocumentHeld(session) ? undefined :
      Promise.race([awaitWorkerOpenReplica(session) ?? Promise.resolve()])
        .then(() => awaitWorkerOpenReplica(session));
    return loaded ? Promise.race([workerRead, loaded.then(() => execute(false))]) : workerRead;
  };
  const round = <T>(
    call: () => Promise<T>, main: () => T | Promise<T>, admission?: RoundAdmission
  ): Promise<T> => {
    admittedRounds += 1;
    return enqueue(async () => {
      if (editorPeer()) await awaitWorkerOpenReplica(session);
      const ready = async () => {
        while (!initialized && !retirementReason) {
          const revision = layoutRevision;
          await Promise.race([hooks.laidOut(), stopped, retired]);
          assertCurrent();
          if (retirementReason) break;
          if (revision !== layoutRevision) continue;
          await initializeNow();
        }
      };
      await ready();
      if (retirementReason) {
        await admission?.flush();
        admission?.validate();
        return main();
      }
      await admission?.flush();
      await ready();
      assertCurrent();
      admission?.validate();
      return retirementReason ? main() : call();
    }).finally(() => { admittedRounds -= 1; });
  };
  const workerSnapshot = (): DocxProposalSnapshot | null => retirementReason ? retiredRegistry!.snapshot() : mirror ? {
    version: peerReady() ? session.version() : mirror.version,
    previewVersion: mirror.proposals.previewVersion,
    proposals: mirror.proposals.entries.map(({ record }) => ({
      ...record, paragraph: { ...record.paragraph }, revisionIds: [...record.revisionIds],
    })),
  } : null;
  const preview = (): ReturnType<typeof proposalRevisionPreview> => {
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
    if (changed.length > 0) hooks.peerUpdated?.(changed);
  };
  const recordCorrespondence = (workerVersion: string, unchanged: boolean, vector: Uint8Array): void => {
    const peerVector = session.encodeStateVector();
    correspondence = unchanged && sameVector(peerVector, vector)
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
    if (pending && !editorPeer() && !peerReady()) session.mirrorWorkerDocument(
      { ...previous, version: previous.version + '~' }
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
      if (hooks.current() && pending && !editorPeer() && workerOpenReplicaPending(session)) {
        session.mirrorWorkerDocument(previous);
      }
      throw error;
    } finally {
      mutating -= 1;
    }
    assertCurrent();
    if (retirementReason) throw new Error('The document worker is unavailable');
    if (
      (!editorPeer() && op.kind === 'setStates' && reply.result?.ok === true) ||
      reply.changedStories.length > 0 ||
      JSON.stringify(previous.proposals) !== JSON.stringify(reply.mirror.proposals)
    ) holdsState = true;
    if (peerReady() && reply.peerDiff) {
      const unchanged = session.version() === postedVersion;
      integrate(reply.peerDiff, [...new Set([...reply.changedStories, ...(reply.projectionStories ?? [])])]);
      recordCorrespondence(reply.mirror.version, unchanged, reply.stateVector);
      holdsState = false;
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
    main: () => Promise<DocxProposalResult>, admission?: RoundAdmission
  ): Promise<DocxProposalResult> => round<DocxProposalResult>(async () => {
    const prepare = () => {
      admission?.validate();
      if (!editorPeer()) return op;
      const version = peerReady() ? session.version() : mirror!.version;
      const token = op.request.expectVersion;
      const retry = op.kind === 'propose' && Array.isArray(op.request.proposals) &&
        op.request.proposals.every((input) => mirror?.proposals.entries.some((entry) =>
          entry.record.id === input.id && entry.key === retryKey(input)));
      if (!retry && token !== version && !(correspondence?.peer === version && correspondence.worker === token)) {
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
  }, main, admission);
  const retiredMutation = (call: () => DocxProposalResult): Promise<DocxProposalResult> => {
    const before = retiredRegistry!.snapshot();
    const since = session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
    const result = call();
    if (result.ok) {
      const changed = session.storiesChangedSince(since).stories;
      const stories = [...new Set([...changed, ...before.proposals, ...result.snapshot.proposals]
        .flatMap((entry) => typeof entry === 'string' ? [entry] : [entry.paragraph.story]))];
      hooks.projectionChanged?.(stories);
      if (stories.length > 0) hooks.peerUpdated?.(stories);
      if (changed.length > 0) hooks.contentChanged();
      if (changed.length > 0 || before.previewVersion !== result.snapshot.previewVersion) hooks.relayout();
    }
    return Promise.resolve(result);
  };
  const initializeNow = async (): Promise<void> => {
    if (retirementReason) return;
    assertCurrent();
    if (initialized || retirementReason) return;
    if (initializationFailure) throw initializationFailure.error;
    const revision = layoutRevision;
    const reply = await Promise.race([worker.proposal({ kind: 'snapshot' }), retired.then(() => null)]).catch((error: unknown) => {
      if (retirementReason) return null;
      if (revision === layoutRevision && (editorRounds || mirror === null)) initializationFailure = { error };
      throw error;
    });
    assertCurrent();
    if (!reply || retirementReason || revision !== layoutRevision) return;
    const previousVersion = mirror?.version;
    initialized = true;
    initializing = Promise.resolve();
    store(reply);
    if (peerReady()) recordCorrespondence(reply.mirror.version, true, reply.stateVector);
    hooks.adopted(reply.mirror.version);
    if (previousVersion !== undefined && previousVersion !== reply.mirror.version) hooks.relayout();
  };
  const authority: RegisteredAuthority = {
    hasEditorRounds: () => editorRounds,
    readsInWorker: () => true,
    get initialized() { return initialized; },
    snapshot: workerSnapshot,
    revisionPreview: preview,
    previewVersion: () => (workerSnapshot()?.previewVersion ?? 0) +
      (retirementReason || (editorPeer() && peerReady()) ? session.getProposals().previewVersion : 0),
    workerCoversPeer: (version) => {
      if (retirementReason || mutating > 0 || correspondence?.peer !== version || session.version() !== version) return false;
      const revision = worker.stateRevision?.();
      return revision?.owner === correspondence.revision?.owner && revision?.sequence === correspondence.revision?.sequence;
    },
    retirementReason: () => retirementReason,
    retire(reason) {
      if (retirementReason) return false;
      if (holdsState) throw new Error('The resident worker holds proposals the main thread cannot rebuild');
      retiredRegistry = session.createWorkerProposalRegistry(mirror?.proposals ?? { previewVersion: 0, entries: [] });
      retiredRegistry.subscribe(notify);
      if (mutating > 0) {
        const error = new Error('The document worker is unavailable');
        for (const reject of pendingCalls) reject(error);
      }
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
    residentOperation: (task) => enqueue(async () => {
      const result = await task();
      assertCurrent();
      return result;
    }),
    restart() {
      if (retirementReason || (!initialized && admittedRounds === 0) || holdsState || failure || !hooks.current()) return;
      initialized = false;
      layoutRevision += 1;
      initializing = null;
      initializationFailure = null;
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
      initializing = Promise.resolve(awaitWorkerOpenReplica(session)).then(() =>
        Promise.race([hooks.laidOut(), stopped, retired])
      ).then(() => enqueue(initializeNow));
      void initializing.catch(() => {});
      return initializing;
    },
    geometry: () => geometry,
    holdsWorkerState: () => !retirementReason && (holdsState || mutating > 0),
    holdsCommittedWorkerState: () => holdsState,
    failure: () => failure?.error,
    fail: (error) => {
      if (failure) return;
      failure = { error };
      failWaiting(error);
      for (const reject of pendingCalls) reject(error);
    },
    propose: (request, main, admission) => mutate({ kind: 'propose', request },
      () => retiredRegistry ? retiredMutation(() => retiredRegistry!.propose(request)) : main(request), admission),
    setStates: (request, main, admission) => mutate({ kind: 'setStates', request },
      () => retiredRegistry ? retiredMutation(() => retiredRegistry!.setStates(request)) : main(request), admission),
    withdraw: (request, main, admission) => mutate({ kind: 'withdraw', request },
      () => retiredRegistry ? retiredMutation(() => retiredRegistry!.withdraw(request)) : main(request), admission),
    removeComment: (id, main) => route(async () => {
      await sendMutation({ kind: 'removeComment', id });
    }, main),
    handedOverRequest: (request) =>
      !retirementReason && correspondence && request.expectVersion === correspondence.worker &&
      session.version() === correspondence.peer
        ? { ...request, expectVersion: correspondence.peer }
        : request,
    getProposals: (main, admission) => round(async () => {
      if (peerReady() && editorPeer()) {
        const version = session.version();
        const reply = await worker.proposal({ kind: 'snapshot' });
        assertCurrent();
        if (!initialized) {
          initialized = true;
          hooks.adopted(reply.mirror.version);
        }
        store(reply);
        recordCorrespondence(reply.mirror.version, session.version() === version, reply.stateVector);
      }
      return workerSnapshot()!;
    }, () => retiredRegistry ? Promise.resolve(retiredRegistry.snapshot()) : main(), admission),
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
  const activated = activationListeners.get(session);
  activationListeners.delete(session);
  activated?.forEach((listener) => listener());
  return authority;
}

export function workerProposalAuthority(session: YrsSession, includeHydrating = false): WorkerProposalAuthority | null {
  const authority = authorities.get(session);
  if (!authority) {
    const ordinary = baseAuthority.workerProposalAuthority(session);
    return ordinary ? extendBaseAuthority(session, ordinary) : null;
  }
  return authority && !authority.retirementReason() && authority.readsInWorker() && workerOpenReplicaPending(session) &&
    (!workerOpenReplicaStarted(session) || authority.holdsWorkerState() || (includeHydrating && authority.initialized))
    ? authority
    : null;
}

const activators = new WeakMap<YrsSession, () => WorkerProposalAuthority | null>();
const activationListeners = new WeakMap<YrsSession, Set<() => void>>();
const extendedBaseAuthorities = new WeakMap<baseAuthority.WorkerProposalAuthority, WorkerProposalAuthority>();

function extendBaseAuthority(session: YrsSession, authority: baseAuthority.WorkerProposalAuthority): WorkerProposalAuthority {
  let extended = extendedBaseAuthorities.get(authority);
  if (!extended) {
    extended = Object.assign(authority, {
      snapshot: () => authority.initialized ? session.getProposals() : null,
      revisionPreview: () => proposalRevisionPreview(session.getProposals()),
      previewVersion: () => session.getProposals().previewVersion,
      workerCoversPeer: () => false,
      retirementReason: () => null,
      retire: () => false,
    });
    extendedBaseAuthorities.set(authority, extended);
  }
  return extended;
}

export function installEditorWorkerProposalActivation(session: YrsSession, activate: () => WorkerProposalAuthority | null): void {
  activators.set(session, activate);
}

export function editorWorkerProposalActivationAvailable(session: YrsSession): boolean {
  return activators.has(session);
}

export async function activateEditorWorkerProposalRounds(
  session: YrsSession, admitted: () => boolean = () => true
): Promise<WorkerProposalAuthority | null> {
  await awaitWorkerOpenReplica(session);
  if (!admitted()) return null;
  return authorities.get(session) ?? activators.get(session)?.() ?? null;
}

export function subscribeEditorWorkerProposalAuthority(session: YrsSession, listener: () => void): () => void {
  let unsubscribe = authorities.get(session)?.subscribe(listener);
  const activated = () => { unsubscribe = authorities.get(session)?.subscribe(listener); listener(); };
  let listeners = activationListeners.get(session);
  if (!listeners) activationListeners.set(session, listeners = new Set());
  if (!unsubscribe) listeners.add(activated);
  return () => { listeners.delete(activated); unsubscribe?.(); };
}

export function workerProposalRoundAuthority(session: YrsSession): WorkerProposalAuthority | null {
  return authorities.get(session) ?? workerProposalAuthority(session);
}

export function registeredWorkerProposalAuthority(session: YrsSession): WorkerProposalAuthority | null {
  const editor = authorities.get(session);
  if (editor) return editor;
  const ordinary = baseAuthority.registeredWorkerProposalAuthority(session);
  return ordinary ? extendBaseAuthority(session, ordinary) : null;
}

export const beginWorkerProposalHandover = baseAuthority.beginWorkerProposalHandover;

export function hasEditorWorkerProposalRounds(session: YrsSession): boolean {
  return authorities.get(session)?.hasEditorRounds() === true;
}

export function workerProposalFailure(session: YrsSession): unknown {
  return authorities.get(session)?.failure() ?? baseAuthority.workerProposalFailure(session);
}

/**
 * Maps a worker token only while the peer still holds its corresponding version.
 */
export function handedOverRequest<T extends { expectVersion: string }>(
  session: YrsSession,
  request: T
): T {
  const authority = authorities.get(session);
  return authority ? workerOpenRequest(session, authority.handedOverRequest(request))
    : baseAuthority.handedOverRequest(session, request);
}

export function snapshotWorkerProposalPeer(session: YrsSession): Promise<PeerSnapshot> | null {
  return authorities.get(session)?.peerSnapshot() ?? null;
}

export function failWorkerProposalAuthority(session: YrsSession, error: unknown): void {
  authorities.get(session)?.fail(error);
  baseAuthority.failWorkerProposalAuthority(session, error);
}

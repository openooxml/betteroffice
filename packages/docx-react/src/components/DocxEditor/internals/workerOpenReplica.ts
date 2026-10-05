import type { YrsSession } from '@betteroffice/docx/yrs';
import { yieldToMainThread } from './yieldToMainThread';

export type WorkerOpenFallbackReason = 'failure' | { syncAccess: string };

type LayoutProgress = 'page' | 'provisional' | 'complete';

interface PendingReplica {
  ready: Promise<void>;
  /** A viewer document held only in its worker: nothing loads it here until the hold is released. */
  held?: true;
  start(): void;
  ensure(reason?: WorkerOpenFallbackReason): void;
  requestReady(): void;
  layoutProgress(progress: LayoutProgress): void;
  fail(error: unknown): void;
  cancel(): void;
  pending: boolean;
  readonly started: boolean;
  readonly hydrated: boolean;
  initialVersion: string;
  hydratingVersion?: string;
  readyVersion?: string;
  loadedVersion?: string;
  /** The worker's version for the opened state, once the session mirrors the worker. */
  mirrorVersion?: string;
  /** The worker's version the replica took over, when it hydrated from worker-held proposals. */
  handoverVersion?: string;
}

const replicas = new WeakMap<YrsSession, PendingReplica>();
const releases = new WeakMap<YrsSession, () => PendingReplica>();

/** A main-thread copy of a viewer document was asked for while the worker holds it. */
export class WorkerOpenDocumentHeldError extends Error {
  constructor(readonly access?: string) {
    super(`${access ?? 'This call'} needs the document on the main thread, which a viewer session does not load`);
    this.name = 'WorkerOpenDocumentHeldError';
  }
}

const LAYOUT_STALL_MS = 3000;
const LAYOUT_WAIT_LIMIT_MS = 30_000;

/**
 * Registers the document of a viewer session as held only in its worker. It reads as a replica still
 * pending, so worker routing and version bookkeeping work as before, but nothing loads it here:
 * starting, awaiting or ensuring it throws {@link WorkerOpenDocumentHeldError}. `release` registers the
 * editor replica with {@link deferWorkerOpenReplica}; {@link releaseWorkerOpenDocument} calls it once
 * the session leaves viewer kind.
 */
export function holdWorkerOpenDocument(session: YrsSession, release: () => PendingReplica): void {
  let reject!: (error: unknown) => void;
  const ready = new Promise<void>((_, no) => {
    reject = no;
  });
  void ready.catch(() => {});
  const held: PendingReplica = {
    ready,
    held: true,
    pending: true,
    started: false,
    hydrated: false,
    requestReady() {},
    layoutProgress() {},
    initialVersion: session.version(),
    start() {},
    ensure(reason = 'failure') {
      throw new WorkerOpenDocumentHeldError(reason === 'failure' ? undefined : reason.syncAccess);
    },
    cancel() {
      held.fail(new Error('The document changed while opening the replica'));
    },
    fail(error) {
      if (!held.pending) return;
      held.pending = false;
      reject(error);
    },
  };
  replicas.set(session, held);
  releases.set(session, release);
}

/** Whether `session` is a viewer document held in its worker; a failed hold still never loads here. */
export function workerOpenDocumentHeld(session: YrsSession): boolean {
  return replicas.get(session)?.held === true;
}

/**
 * Ends the hold on `session`: registers its editor replica, carrying the versions the hold recorded,
 * and returns it unstarted. Undefined when `session` is not held (or its hold failed).
 */
export function releaseWorkerOpenDocument(session: YrsSession): PendingReplica | undefined {
  const held = replicas.get(session);
  const release = releases.get(session);
  if (!held?.held || !held.pending || !release) return undefined;
  releases.delete(session);
  const replica = release();
  replica.initialVersion = held.initialVersion;
  replica.mirrorVersion = held.mirrorVersion;
  replica.handoverVersion = held.handoverVersion;
  return replica;
}

export function deferWorkerOpenReplica(
  session: YrsSession,
  hydrate: () => Promise<(() => void) | readonly (() => void)[]>,
  fallback: (reason: WorkerOpenFallbackReason) => void,
  onReady: () => void,
  lifecycle?: { current(): boolean; cancel(): void; waitForLayout?: boolean }
): PendingReplica {
  let resolve!: () => void;
  let reject!: (error: unknown) => void;
  let started = false;
  let finishing = false;
  let hydrated = false;
  let layoutComplete = false;
  let readinessRequested = false;
  let stallTimer: ReturnType<typeof setTimeout> | null = null;
  let limitTimer: ReturnType<typeof setTimeout> | null = null;
  let failure: unknown;
  let steps: readonly (() => void)[] | null = null;
  let nextStep = 0;
  const controller = new AbortController();
  const ready = new Promise<void>((yes, no) => {
    resolve = yes;
    reject = no;
  });
  void ready.catch(() => {});
  const current = (): boolean => {
    if (!replica.pending || controller.signal.aborted) return false;
    if (lifecycle?.current() !== false) return true;
    replica.cancel();
    return false;
  };
  const clearLayoutWait = (): void => {
    if (stallTimer !== null) clearTimeout(stallTimer);
    if (limitTimer !== null) clearTimeout(limitTimer);
    stallTimer = null;
    limitTimer = null;
  };
  const commit = (): void => {
    if (!hydrated || !current()) return;
    clearLayoutWait();
    replica.pending = false;
    controller.abort();
    try {
      onReady();
      resolve();
    } catch (error) {
      failure = error;
      reject(error);
    }
  };
  const resetStallTimer = (): void => {
    if (stallTimer !== null) clearTimeout(stallTimer);
    stallTimer = setTimeout(commit, LAYOUT_STALL_MS);
  };
  const finish = (
    load: (reason: WorkerOpenFallbackReason) => void,
    handoff = false,
    force = false,
    reason: WorkerOpenFallbackReason = 'failure'
  ): void => {
    if (!current() || finishing) return;
    finishing = true;
    try {
      try {
        load(reason);
        if (!current()) return;
        if (handoff) replica.readyVersion = session.version();
      } catch (error) {
        if (!handoff) throw error;
        if (!current()) return;
        fallback('failure');
        handoff = false;
      }
      if (!current()) return;
      replica.loadedVersion = session.version();
      hydrated = true;
      steps = null;
      if (lifecycle?.waitForLayout && handoff && !force && !readinessRequested && !layoutComplete) {
        resetStallTimer();
        limitTimer = setTimeout(commit, LAYOUT_WAIT_LIMIT_MS);
      } else commit();
    } catch (error) {
      failure = error;
      replica.pending = false;
      steps = null;
      clearLayoutWait();
      controller.abort();
      reject(error);
    } finally {
      finishing = false;
    }
  };
  const loadRemaining = (): void => {
    while (steps && nextStep < steps.length && current()) steps[nextStep++]!();
  };
  const hydrateInTasks = async (load: (() => void) | readonly (() => void)[]): Promise<void> => {
    if (!current()) return;
    const plan = typeof load === 'function' ? [load] : load;
    steps = plan;
    try {
      while (nextStep < plan.length) {
        if (!current()) return;
        finishing = true;
        try {
          plan[nextStep++]!();
          if (current()) replica.hydratingVersion = session.version();
        } finally {
          finishing = false;
        }
        if (!current()) return;
        if (nextStep < plan.length) {
          await yieldToMainThread();
          if (!current()) return;
        }
      }
      finish(() => {}, true);
    } catch (error) {
      finish(() => { throw error; }, true);
    }
  };
  const replica: PendingReplica = {
    ready,
    pending: true,
    get started() {
      return started || !replica.pending;
    },
    get hydrated() {
      return hydrated;
    },
    initialVersion: session.version(),
    start() {
      if (started || !replica.pending) return;
      started = true;
      void hydrate().then(
        hydrateInTasks,
        () => finish(fallback)
      );
    },
    ensure(reason = 'failure') {
      finish(hydrated ? () => {} : steps ? loadRemaining : fallback, steps !== null, true, reason);
      if (failure !== undefined) throw failure;
    },
    requestReady() {
      readinessRequested = true;
      commit();
    },
    layoutProgress(progress) {
      if (!current()) return;
      if (progress !== 'page') layoutComplete = progress === 'complete';
      if (progress === 'complete') commit();
      else if (hydrated && stallTimer !== null) resetStallTimer();
    },
    cancel() {
      if (!replica.pending) return;
      replica.fail(new Error('The document changed while opening the replica'));
      lifecycle?.cancel();
    },
    fail(error) {
      if (!replica.pending) return;
      replica.pending = false;
      steps = null;
      clearLayoutWait();
      controller.abort();
      failure = error;
      reject(error);
    },
  };
  replicas.set(session, replica);
  return replica;
}

export function workerOpenReplicaStarted(session: YrsSession): boolean {
  return replicas.get(session)?.started === true;
}

export function workerOpenReplicaPending(session: YrsSession): boolean {
  return replicas.get(session)?.pending === true;
}

export function workerOpenReplicaHydrating(session: YrsSession): boolean {
  const replica = replicas.get(session);
  return replica?.pending === true && replica.started && !replica.hydrated;
}

/**
 * Starts loading the replica of `session` from its worker, once, and returns the promise that
 * settles when it is ready; undefined when `session` has no worker-open replica.
 */
export function requestWorkerOpenReplica(session: YrsSession): Promise<void> | undefined {
  const replica = replicas.get(session);
  if (replica?.held && replica.pending) return Promise.reject(new WorkerOpenDocumentHeldError());
  replica?.start();
  return replica?.ready;
}

export function awaitWorkerOpenReplica(session: YrsSession): Promise<void> | undefined {
  const replica = replicas.get(session);
  if (replica?.held && replica.pending) return Promise.reject(new WorkerOpenDocumentHeldError());
  return replica?.ready;
}

export function requestWorkerOpenReplicaReadiness(session: YrsSession): void {
  replicas.get(session)?.requestReady();
}

export function notifyWorkerOpenLayoutProgress(session: YrsSession, progress: LayoutProgress): void {
  replicas.get(session)?.layoutProgress(progress);
}

/** The version `session` had when its replica loaded; a later version holds a newer change. */
export function workerOpenReplicaLoadedVersion(session: YrsSession): string | undefined {
  return replicas.get(session)?.loadedVersion;
}

export function ensureWorkerOpenReplica(session: YrsSession, syncAccess?: string): void {
  replicas.get(session)?.ensure(syncAccess === undefined ? 'failure' : { syncAccess });
}

export function failWorkerOpenReplica(session: YrsSession, error: unknown): void {
  replicas.get(session)?.fail(error);
}

export function workerOpenSourceVersion(session: YrsSession, version: string | null): string | null {
  const replica = replicas.get(session);
  if (!replica || version === null) return version;
  let mapped = version;
  if (replica.pending && mapped === replica.hydratingVersion) {
    mapped = replica.handoverVersion ?? replica.mirrorVersion ?? replica.initialVersion;
  }
  if (mapped === replica.initialVersion) {
    mapped = replica.mirrorVersion ?? replica.readyVersion ?? mapped;
  }
  if (mapped === replica.handoverVersion && replica.readyVersion !== undefined) {
    mapped = replica.readyVersion;
  }
  return mapped;
}

export function workerOpenRequest<T extends { expectVersion: string }>(session: YrsSession, request: T): T {
  const replica = replicas.get(session);
  return replica && !replica.pending && replica.readyVersion !== undefined &&
    request.expectVersion === replica.handoverVersion && session.version() === replica.readyVersion
    ? { ...request, expectVersion: replica.readyVersion }
    : request;
}

/** Layouts of the worker's `version` show the state the replica hydrates with. */
export function adoptWorkerOpenHandoverVersion(session: YrsSession, version: string): void {
  const replica = replicas.get(session);
  if (replica?.pending) replica.handoverVersion = version;
}

/** Layouts of the opened state now carry `version`, the worker's version the session mirrors. */
export function adoptWorkerOpenMirrorVersion(session: YrsSession, version: string): void {
  const replica = replicas.get(session);
  if (replica?.pending) replica.mirrorVersion = version;
}

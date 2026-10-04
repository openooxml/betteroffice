import type { YrsSession } from '@betteroffice/docx/yrs';

export type WorkerOpenFallbackReason = 'failure' | { syncAccess: string };

interface PendingReplica {
  ready: Promise<void>;
  start(): void;
  ensure(reason?: WorkerOpenFallbackReason): void;
  fail(error: unknown): void;
  cancel(): void;
  pending: boolean;
  onDemand?: WorkerOpenReplicaDemand;
  readonly started: boolean;
  initialVersion: string;
  readyVersion?: string;
  loadedVersion?: string;
  /** The worker's version for the opened state, once the session mirrors the worker. */
  mirrorVersion?: string;
  /** The worker's version the replica took over, when it hydrated from worker-held proposals. */
  handoverVersion?: string;
}

/** Loads a replica only when asked: `request` starts it at the point its owner allows. */
export interface WorkerOpenReplicaDemand {
  active(): boolean;
  request(): void;
}

const replicas = new WeakMap<YrsSession, PendingReplica>();

export function deferWorkerOpenReplica(
  session: YrsSession,
  hydrate: () => Promise<() => void>,
  fallback: (reason: WorkerOpenFallbackReason) => void,
  onReady: () => void,
  onDemand?: WorkerOpenReplicaDemand
): PendingReplica {
  let resolve!: () => void;
  let reject!: (error: unknown) => void;
  let started = false;
  let finishing = false;
  let failure: unknown;
  const ready = new Promise<void>((yes, no) => {
    resolve = yes;
    reject = no;
  });
  void ready.catch(() => {});
  const finish = (
    load: (reason: WorkerOpenFallbackReason) => void,
    handoff = false,
    reason: WorkerOpenFallbackReason = 'failure'
  ): void => {
    if (!replica.pending || finishing) return;
    finishing = true;
    try {
      try {
        load(reason);
        if (handoff) replica.readyVersion = session.version();
      } catch (error) {
        if (!handoff) throw error;
        fallback('failure');
      }
      replica.loadedVersion = session.version();
      replica.pending = false;
      onReady();
      resolve();
    } catch (error) {
      failure = error;
      replica.pending = false;
      reject(error);
    } finally {
      finishing = false;
    }
  };
  const replica: PendingReplica = {
    ready,
    pending: true,
    onDemand,
    get started() {
      return started || !replica.pending;
    },
    initialVersion: session.version(),
    start() {
      if (started || !replica.pending) return;
      started = true;
      void hydrate().then(
        (load) => finish(load, true),
        () => finish(fallback)
      );
    },
    ensure(reason = 'failure') {
      finish(fallback, false, reason);
      if (failure !== undefined) throw failure;
    },
    cancel() {
      replica.fail(new Error('The document changed while opening the replica'));
    },
    fail(error) {
      if (!replica.pending) return;
      replica.pending = false;
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

/**
 * Starts loading the replica of `session` from its worker, once, and returns the promise that
 * settles when it is ready; undefined when `session` has no worker-open replica.
 */
export function requestWorkerOpenReplica(session: YrsSession): Promise<void> | undefined {
  const replica = replicas.get(session);
  replica?.start();
  return replica?.ready;
}

export function awaitWorkerOpenReplica(session: YrsSession, options?: { passive: boolean }): Promise<void> | undefined {
  const replica = replicas.get(session);
  if (!options?.passive && replica?.pending && replica.onDemand?.active() === true) replica.onDemand.request();
  return replica?.ready;
}

export function workerOpenReplicaOnDemand(session: YrsSession): boolean {
  const replica = replicas.get(session);
  return replica?.pending === true && replica.onDemand?.active() === true;
}

/** Asks an on-demand replica of `session` to load; see {@link WorkerOpenReplicaDemand}. */
export function requestOnDemandWorkerOpenReplica(session: YrsSession): void {
  const replica = replicas.get(session);
  if (replica?.pending && replica.onDemand?.active() === true) replica.onDemand.request();
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
  if (mapped === replica.initialVersion) {
    mapped = replica.mirrorVersion ?? replica.readyVersion ?? mapped;
  }
  if (mapped === replica.handoverVersion && replica.readyVersion !== undefined) {
    mapped = replica.readyVersion;
  }
  return mapped;
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

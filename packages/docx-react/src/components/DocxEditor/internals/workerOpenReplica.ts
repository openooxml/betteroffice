import type { YrsSession } from '@betteroffice/docx/yrs';
import { yieldToMainThread } from './yieldToMainThread';

interface PendingReplica {
  ready: Promise<void>;
  start(): void;
  ensure(): void;
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
  hydrate: () => Promise<(() => void) | readonly (() => void)[]>,
  fallback: () => void,
  onReady: () => void,
  onDemand?: WorkerOpenReplicaDemand,
  lifecycle?: { current(): boolean; cancel(): void }
): PendingReplica {
  let resolve!: () => void;
  let reject!: (error: unknown) => void;
  let started = false;
  let finishing = false;
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
  const finish = (load: () => void, handoff = false): void => {
    if (!current() || finishing) return;
    finishing = true;
    try {
      try {
        load();
        if (!current()) return;
        if (handoff) replica.readyVersion = session.version();
      } catch (error) {
        if (!handoff) throw error;
        if (!current()) return;
        fallback();
      }
      if (!current()) return;
      replica.loadedVersion = session.version();
      replica.pending = false;
      steps = null;
      controller.abort();
      onReady();
      resolve();
    } catch (error) {
      failure = error;
      replica.pending = false;
      steps = null;
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
    onDemand,
    get started() {
      return started || !replica.pending;
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
    ensure() {
      finish(steps ? loadRemaining : fallback, steps !== null);
      if (failure !== undefined) throw failure;
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

/**
 * Starts loading the replica of `session` from its worker, once, and returns the promise that
 * settles when it is ready; undefined when `session` has no worker-open replica.
 */
export function requestWorkerOpenReplica(session: YrsSession): Promise<void> | undefined {
  const replica = replicas.get(session);
  replica?.start();
  return replica?.ready;
}

export function awaitWorkerOpenReplica(session: YrsSession): Promise<void> | undefined {
  const replica = replicas.get(session);
  if (replica?.pending && replica.onDemand?.active() === true) replica.onDemand.request();
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

export function ensureWorkerOpenReplica(session: YrsSession): void {
  replicas.get(session)?.ensure();
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

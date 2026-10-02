import type { Comment } from '@betteroffice/docx/types/content';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { workerOpenReplicaPending, workerOpenReplicaStarted } from './workerOpenReplica';

/** Saves the document in its worker; null when the host has no document to save. */
export type WorkerOpenSave = (
  comments: Comment[]
) => Promise<{ bytes: ArrayBuffer; full: boolean }> | null;

const savers = new WeakMap<YrsSession, WorkerOpenSave>();
const savedOriginals = new WeakMap<YrsSession, ArrayBuffer>();
const saving = new WeakMap<YrsSession, Set<Promise<unknown>>>();

export function registerWorkerOpenSave(session: YrsSession, save: WorkerOpenSave): void {
  savers.set(session, save);
}

/** The worker's save of `session` while its replica has not started loading. */
export function workerOpenSave(session: YrsSession): WorkerOpenSave | null {
  const save = savers.get(session);
  return save && workerOpenReplicaPending(session) && !workerOpenReplicaStarted(session)
    ? save
    : null;
}

/** `bytes`, the worker's last save of `session`, is the package the replica's next save starts from. */
export function recordWorkerOpenSave(session: YrsSession, bytes: ArrayBuffer): void {
  savedOriginals.set(session, bytes);
}

export function trackWorkerOpenSave<T>(session: YrsSession, pending: Promise<T>): Promise<T> {
  const saves = saving.get(session) ?? new Set<Promise<unknown>>();
  saving.set(session, saves);
  const tracked = pending.finally(() => { saves.delete(tracked); });
  saves.add(tracked);
  return tracked;
}

export function awaitWorkerOpenSaves(session: YrsSession): Promise<void> | undefined {
  const saves = saving.get(session);
  if (saves?.size) return Promise.allSettled(saves).then(() => {});
}

export function peekWorkerOpenSave(session: YrsSession): ArrayBuffer | undefined {
  return savedOriginals.get(session);
}

export function takeWorkerOpenSave(session: YrsSession): ArrayBuffer | undefined {
  const bytes = savedOriginals.get(session);
  savedOriginals.delete(session);
  return bytes;
}

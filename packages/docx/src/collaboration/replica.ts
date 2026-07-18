import type { YrsSelection, YrsSession } from '../yrs';
import type { CollaborationReplica } from './types';

export interface WorkerCollaborationTarget {
  invalidate(update: Uint8Array, selection: YrsSelection | null): void;
}

export function createDirectCollaborationReplica(session: YrsSession): CollaborationReplica {
  return session;
}

export function createWorkerCollaborationReplica(
  session: YrsSession,
  worker: WorkerCollaborationTarget
): CollaborationReplica {
  return {
    clientId: session.clientId,
    encodeStateVector: () => session.encodeStateVector(),
    encodeStateAsUpdate: (remoteStateVector) =>
      session.encodeStateAsUpdate(remoteStateVector?.slice()),
    applyUpdate: (update) => {
      const owned = update.slice();
      session.applyUpdate(owned);
      worker.invalidate(owned, session.selection());
    },
    onUpdate: (listener) => session.onUpdate(listener),
  };
}

import type { AsyncCollaborationReplica, CollaborationReplica } from './types';

/** @experimental */
export interface AsyncReplicaAdapter extends CollaborationReplica {
  readonly async: AsyncCollaborationReplica;
}

const adapters = new WeakMap<CollaborationReplica, AsyncCollaborationReplica>();

export function asyncReplicaFor(replica: CollaborationReplica): AsyncCollaborationReplica | undefined {
  return adapters.get(replica);
}

/** @experimental Wraps a worker replica so `CollaborationProvider` and `onReplica` accept it unchanged. */
export function createAsyncCollaborationReplica(replica: AsyncCollaborationReplica): AsyncReplicaAdapter {
  const unavailable = (): never => { throw new Error('Worker replica encoding requires the async adapter'); };
  const adapter: AsyncReplicaAdapter = {
    clientId: replica.clientId, async: replica,
    encodeStateVector: unavailable, encodeStateAsUpdate: unavailable,
    applyUpdate: (update) => replica.applyUpdate(update.slice()),
    onUpdate: (listener) => replica.onUpdate(listener),
  };
  adapters.set(adapter, replica);
  return adapter;
}

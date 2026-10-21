import type { CollaborationReplica, CollaborationUpdateOrigin } from './types';

/** @experimental */
export interface AsyncCollaborationReplica {
  readonly clientId: number;
  encodeStateVector(): Promise<Uint8Array>;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Promise<Uint8Array>;
  applyUpdate(update: Uint8Array): Promise<unknown>;
  /** Applies updates in order as one merge; rejects at the first refused update, keeping the ones before it. */
  applyUpdateChunks?(updates: readonly Uint8Array[]): Promise<unknown>;
  onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
}

/** @experimental */
export interface AsyncReplicaAdapter extends CollaborationReplica {
  readonly async: AsyncCollaborationReplica;
}

const adapters = new WeakMap<CollaborationReplica, AsyncCollaborationReplica>();

export function asyncReplicaFor(replica: CollaborationReplica): AsyncCollaborationReplica | undefined {
  return adapters.get(replica);
}

/** @experimental */
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

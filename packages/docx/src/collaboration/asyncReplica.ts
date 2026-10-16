import type { AsyncCollaborationReplica, CollaborationReplica } from './types';

/**
 * @experimental A `CollaborationReplica` whose state lives in a worker. `CollaborationProvider` reads it
 * through `async`. The synchronous `encodeStateVector` and `encodeStateAsUpdate` throw; use
 * `async.encodeStateVector()` and `async.encodeStateAsUpdate()`, or open the editor with `legacySyncApi`
 * for a synchronous replica.
 */
export interface AsyncReplicaAdapter extends CollaborationReplica {
  readonly async: AsyncCollaborationReplica;
}

const adapters = new WeakMap<CollaborationReplica, AsyncCollaborationReplica>();

export function asyncReplicaFor(replica: CollaborationReplica): AsyncCollaborationReplica | undefined {
  return adapters.get(replica);
}

function unavailable(method: 'encodeStateVector' | 'encodeStateAsUpdate'): () => never {
  return () => {
    throw new Error(
      `${method} is not available on a worker-hosted replica; use async.${method}() or open the editor with legacySyncApi for synchronous replica access`
    );
  };
}

/** @experimental Wraps a worker replica so `CollaborationProvider` and `onReplica` accept it unchanged. */
export function createAsyncCollaborationReplica(replica: AsyncCollaborationReplica): AsyncReplicaAdapter {
  const adapter: AsyncReplicaAdapter = {
    clientId: replica.clientId, async: replica,
    encodeStateVector: unavailable('encodeStateVector'),
    encodeStateAsUpdate: unavailable('encodeStateAsUpdate'),
    applyUpdate: (update) => replica.applyUpdate(update.slice()),
    onUpdate: (listener) => replica.onUpdate(listener),
  };
  adapters.set(adapter, replica);
  return adapter;
}

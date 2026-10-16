import type { AsyncCollaborationReplica, CollaborationReplica } from './types';

/**
 * @experimental A `CollaborationReplica` whose state lives in a worker. `CollaborationProvider` reads it
 * through `async`. The synchronous `encodeStateVector`, `encodeStateAsUpdate` and `encodeStateAsUpdateChunks`
 * throw; use their `async` counterparts, or open the editor with `legacySyncApi` for a synchronous replica.
 * `applyUpdate` and `async`'s apply methods copy the caller's updates, which stay usable.
 */
export interface AsyncReplicaAdapter extends CollaborationReplica {
  readonly async: AsyncCollaborationReplica;
}

const adapters = new WeakMap<CollaborationReplica, AsyncCollaborationReplica>();

export function asyncReplicaFor(replica: CollaborationReplica): AsyncCollaborationReplica | undefined {
  return adapters.get(replica);
}

function unavailable(method: 'encodeStateVector' | 'encodeStateAsUpdate' | 'encodeStateAsUpdateChunks'): () => never {
  return () => {
    throw new Error(
      `${method} is not available on a worker-hosted replica; use async.${method}() or open the editor with legacySyncApi for synchronous replica access`
    );
  };
}

/** @experimental Wraps a worker replica so `CollaborationProvider` and `onReplica` accept it unchanged. */
export function createAsyncCollaborationReplica(replica: AsyncCollaborationReplica): AsyncReplicaAdapter {
  const hosted: AsyncCollaborationReplica = {
    clientId: replica.clientId,
    encodeStateVector: () => replica.encodeStateVector(),
    encodeStateAsUpdate: (remoteStateVector) => replica.encodeStateAsUpdate(remoteStateVector),
    ...(replica.encodeStateAsUpdateChunks && {
      encodeStateAsUpdateChunks: (remoteStateVector: Uint8Array, maxChunkBytes: number) =>
        replica.encodeStateAsUpdateChunks!(remoteStateVector, maxChunkBytes),
    }),
    applyUpdate: (update) => replica.applyUpdate(update.slice()),
    ...(replica.applyUpdateChunks && {
      applyUpdateChunks: (updates: readonly Uint8Array[]) => replica.applyUpdateChunks!(updates.map((update) => update.slice())),
    }),
    onUpdate: (listener) => replica.onUpdate(listener),
  };
  const adapter: AsyncReplicaAdapter = {
    clientId: replica.clientId, async: hosted,
    encodeStateVector: unavailable('encodeStateVector'),
    encodeStateAsUpdate: unavailable('encodeStateAsUpdate'),
    ...(hosted.encodeStateAsUpdateChunks && { encodeStateAsUpdateChunks: unavailable('encodeStateAsUpdateChunks') }),
    applyUpdate: hosted.applyUpdate,
    onUpdate: hosted.onUpdate,
  };
  adapters.set(adapter, replica);
  return adapter;
}

// @betteroffice/docx declaration collaboration/asyncReplica.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { AsyncCollaborationReplica, CollaborationReplica } from './types';
import type { CollaborationPipeChannel } from '../../../../shared/collaboration-pipe/main';
/**
 * @experimental A `CollaborationReplica` whose state lives in a worker. `CollaborationProvider` reads it
 * through `async`. The synchronous `encodeStateVector`, `encodeStateAsUpdate` and `encodeStateAsUpdateChunks`
 * throw; use their `async` counterparts, or open the editor with `legacySyncApi` for a synchronous replica.
 * `applyUpdate` and `async`'s apply methods copy the caller's updates, which stay usable.
 */
export interface AsyncReplicaAdapter extends CollaborationReplica {
    readonly async: AsyncCollaborationReplica;
}
export declare function asyncReplicaFor(replica: CollaborationReplica): AsyncCollaborationReplica | undefined;
/**
 * @internal Lets a `CollaborationProvider` built on `replica` run its protocol in the worker that hosts the
 * replica, through `channel`; the provider on the calling thread then only moves transport bytes.
 */
export declare function attachCollaborationPipe(replica: AsyncReplicaAdapter, channel: CollaborationPipeChannel): void;
export declare function collaborationPipeFor(replica: CollaborationReplica): CollaborationPipeChannel | undefined;
/** @experimental Wraps a worker replica so `CollaborationProvider` and `onReplica` accept it unchanged. */
export declare function createAsyncCollaborationReplica(replica: AsyncCollaborationReplica): AsyncReplicaAdapter;

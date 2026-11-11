// @betteroffice/docx declaration shared/collaboration-pipe/worker.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type PipeCommandPayload, type PipeCursor, type PipeError, type PipeOutput, type PipeOutputBatch, type PipePeer, type PipeProviderOptions } from './protocol';
type UpdateOrigin = 'local' | 'remote';
type TransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: Uint8Array;
} | {
    type: 'close';
    reason?: string;
} | {
    type: 'error';
    error: unknown;
};
/** The synchronous replica the worker's provider runs against; merges wait for the end of the burst. */
export interface PipeReplica {
    readonly clientId: number;
    encodeStateVector(): Uint8Array;
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
    encodeStateAsUpdateChunks?(remoteStateVector: Uint8Array, maxChunkBytes: number): Uint8Array[];
    applyUpdate(update: Uint8Array): unknown;
    applyUpdateChunks(updates: readonly Uint8Array[]): unknown;
    onUpdate(listener: (update: Uint8Array, origin: UpdateOrigin) => void): () => void;
}
export interface PipeTransport {
    connect(): void;
    disconnect(): void;
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: TransportEvent) => void): () => void;
}
export interface PipeProvider<C = PipeCursor, P = PipePeer> {
    connect(): void;
    disconnect(): void;
    destroy(): void;
    setCursor(cursor: C | null, broadcast?: boolean): void;
    onStatus(listener: (change: {
        status: Extract<PipeOutput, {
            do: 'status';
        }>['status'];
        synced: boolean;
    }) => void): () => void;
    onError(listener: (error: {
        code: string;
        message: string;
    }) => void): () => void;
    onPeers(listener: (peers: readonly P[]) => void): () => void;
}
/** A merge the provider handed over: settled once the burst's updates merge, failed when they do not. */
export interface PipeArrival {
    settle(applied: unknown): void;
    fail(cause: unknown): void;
}
export type PipeMergeResult = {
    applied: unknown;
} | {
    failed: unknown;
};
export interface CollaborationPipeHostOptions<Prov extends PipeProvider<C, P>, C = PipeCursor, P = PipePeer> {
    clientId: number;
    read: Pick<PipeReplica, 'encodeStateVector' | 'encodeStateAsUpdate' | 'encodeStateAsUpdateChunks'>;
    onUpdate: PipeReplica['onUpdate'];
    /**
     * Merges a burst's updates in one change: `single` is one peer update, which may report a typing inference
     * as `applied`. Returns `failed` for updates the replica refused without changing; throws when it cannot go on.
     */
    merge(updates: Uint8Array[], single: boolean): PipeMergeResult;
    createProvider(replica: PipeReplica, transport: PipeTransport, options: PipeProviderOptions): Prov;
    arrive(provider: Prov): PipeArrival;
    /** The awareness leave `provider` would send now, by the clock it carries; null when it would send none. */
    leave(provider: Prov): {
        clock: number;
        frame: Uint8Array;
    } | null;
    /** The error object a transport `error` event carries for `error`. */
    transportError(error: PipeError): unknown;
    /** Output became pending outside `run`, from the providers' timers. */
    onOutput?(): void;
}
export interface CollaborationPipeHost<C = PipeCursor, P = PipePeer> {
    /** Runs one burst of ops: transport events in order, then one merge of every update they carried. */
    run(payload: PipeCommandPayload<C>): void;
    /** Pending output, in order, and the buffers it holds. */
    drain(): {
        batches: PipeOutputBatch<P>[];
        transfer: ArrayBuffer[];
    } | undefined;
    /** Forgets every provider without sending anything. */
    dispose(): void;
}
export declare function createCollaborationPipeHost<Prov extends PipeProvider<C, P>, C = PipeCursor, P = PipePeer>(options: CollaborationPipeHostOptions<Prov, C, P>): CollaborationPipeHost<C, P>;
export {};

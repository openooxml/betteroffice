// @betteroffice/docx declaration shared/collaboration-pipe/main.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type PipeCursor, type PipeErrorCode, type PipeOp, type PipeOutput, type PipePeer, type PipeProviderOptions } from './protocol';
type Status = 'disconnected' | 'connecting' | 'connected' | 'destroyed';
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
} | {
    type: 'drain';
};
export interface PipeHostTransport {
    connect(): void | Promise<void>;
    disconnect(): void | Promise<void>;
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: TransportEvent) => void): () => void;
}
/** One provider's line to the worker that runs it. */
export interface CollaborationPipeEndpoint<C = PipeCursor> {
    /** Runs `ops` in the worker, in order after earlier posts; settles once the worker has run them. */
    post(ops: PipeOp<C>[], transfer: ArrayBuffer[]): Promise<unknown>;
    close(): void;
}
/** What a worker replica offers a provider that can run in its worker. */
export interface CollaborationPipeChannel<C = PipeCursor, P = PipePeer> {
    /** `receive` takes the worker's output for this provider; `lost` reports the worker is gone for it. */
    open(receive: (out: readonly PipeOutput<P>[]) => void, lost: (reason: string) => void): CollaborationPipeEndpoint<C>;
}
export interface PipedCollaborationOptions<E, P = PipePeer> {
    worker: PipeProviderOptions;
    maxPendingBytes: number;
    hasUser: boolean;
    /** Moves a worker peer's timestamps onto this context's clock; omitted when they are wall-clock times. */
    rebasePeer?(peer: P, offset: number): P;
    makeError(code: PipeErrorCode, message: string, cause?: unknown): E;
    /** `cause` itself when it already is an `E`, else a new `E` with `context` and the cause's message. */
    normalize(code: PipeErrorCode, context: string, cause: unknown): E;
}
/**
 * The main-thread side of a provider whose protocol runs in the document worker. It owns the host transport's
 * lifecycle as a provider does, forwards the transport's events with frame bytes untouched, one post per burst,
 * and sends what the worker answers. Worker output older than main's latest lifecycle call is dropped.
 */
export declare class PipedCollaboration<E extends Error, C = PipeCursor, P = PipePeer> {
    private readonly transport;
    private readonly options;
    private readonly endpoint;
    private readonly statusListeners;
    private readonly errorListeners;
    private readonly peerListeners;
    private nextListenerId;
    private connectionStatus;
    private isSynced;
    private peerList;
    private seq;
    /** The op of main's latest lifecycle call; older worker output describes a connection main has left. */
    private floor;
    /** The op of the transport's latest open or close; sends, status and peers from before it describe a connection the host ended. */
    private sendFloor;
    /** Received frame bytes queued or in flight to the worker. */
    private inboundBytes;
    private flushingFrames;
    private queue;
    private posting;
    private flushQueued;
    private heldTimer;
    /** The queue holds an op that should not wait for held updates. */
    private urgent;
    private unsubscribeTransport;
    private subscription;
    private wantsConnection;
    private connectAttempt;
    private leaveFrame;
    private readonly pendingFrames;
    private queuedBytes;
    private isDestroyed;
    private closed;
    constructor(channel: CollaborationPipeChannel<C, P>, transport: PipeHostTransport, options: PipedCollaborationOptions<E, P>);
    get status(): Status;
    get synced(): boolean;
    get pendingBytes(): number;
    get peers(): readonly P[];
    setCursor(cursor: C | null, broadcast: boolean): void;
    onPeers(listener: (peers: readonly P[]) => void): () => void;
    onStatus(listener: (change: {
        status: Status;
        synced: boolean;
    }) => void): () => void;
    onError(listener: (error: E) => void): () => void;
    connect(): void;
    disconnect(): void;
    destroy(): void;
    private transportError;
    /** Leaves the connection main holds: drops the subscription, disconnects, clears queued bytes and peers. */
    private release;
    /** Fails main's connection with `error`, and the worker's with it. */
    private fail;
    private sendLeave;
    private enqueue;
    private schedule;
    private flush;
    private forward;
    private receive;
    private send;
    private queueFrame;
    private flushPending;
    private clearQueue;
    /**
     * The worker no longer runs this provider. Hands the host transport back at once; reports the loss a task later,
     * so a host that destroys the provider on `onReplica(null)` hears nothing.
     */
    private lose;
    private close;
    private clearPeers;
    private setStatus;
    private emitPeers;
    private report;
}
export {};

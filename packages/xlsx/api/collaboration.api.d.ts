// @betteroffice/xlsx "./collaboration" public API surface
// Regenerate: bun scripts/api-surface.ts --write xlsx
// ---- collaboration/asyncReplica.d.ts ----
import type { CollaborationPipeChannel } from '../../../../shared/collaboration-pipe/main';
import type { AwarenessCursor, AwarenessPeer, CollaborationReplica, CollaborationUpdateOrigin } from './types';
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
export declare function asyncReplicaFor(replica: CollaborationReplica): AsyncCollaborationReplica | undefined;
/** @internal */
export declare function attachCollaborationPipe(adapter: AsyncReplicaAdapter, channel: CollaborationPipeChannel<AwarenessCursor, AwarenessPeer> | undefined): void;
/** @internal */
export declare function collaborationPipeFor(replica: CollaborationReplica): CollaborationPipeChannel<AwarenessCursor, AwarenessPeer> | undefined;
/** @experimental */
export declare function createAsyncCollaborationReplica(replica: AsyncCollaborationReplica): AsyncReplicaAdapter;

// ---- collaboration/awareness.d.ts ----
import type { AwarenessCursor, AwarenessPayload, AwarenessPeer, AwarenessUpdate, CollaborationUser, CollaborationUserOptions } from './types';
export declare const AWARENESS_BROADCAST_INTERVAL_MS = 80;
export declare const AWARENESS_HEARTBEAT_INTERVAL_MS = 20000;
export declare const AWARENESS_PEER_TIMEOUT_MS = 45000;
export declare const AWARENESS_LABEL_DURATION_MS = 3000;
export declare const XLSX_MAX_ROWS = 1048576;
export declare const XLSX_MAX_COLUMNS = 16384;
export declare const MAX_AWARENESS_NAME_LENGTH = 128;
export declare const MAX_AWARENESS_SHEET_LENGTH = 256;
export declare const MAX_TRACKED_AWARENESS_PEERS = 1024;
export declare const AWARENESS_COLORS: readonly ["#0B57D0", "#B3261E", "#146C2E", "#6B4EFF", "#8E24AA", "#00639B", "#8E4E00", "#7D5260"];
export interface AwarenessPeerEntry {
    clock: number;
    state: AwarenessPayload | null;
    lastSeen: number;
    cursorMovedAt: number;
}
export type AwarenessPeerStore = Map<number, AwarenessPeerEntry>;
export interface ApplyAwarenessUpdateOptions {
    maxPeers?: number;
    onPeersDiscarded?: (count: number) => void;
}
export interface ResolvedAwarenessCursor {
    sheetIndex: number;
    range: {
        top: number;
        left: number;
        bottom: number;
        right: number;
    };
}
export declare function colorForClientId(clientId: number): string;
export declare function normalizeAwarenessColor(color: unknown, clientId: number): string;
export declare function normalizeCollaborationUser(user: CollaborationUserOptions | undefined, clientId: number): CollaborationUser;
export declare function applyAwarenessUpdates(store: AwarenessPeerStore, updates: readonly AwarenessUpdate[], localClientId: number, now: number, options?: ApplyAwarenessUpdateOptions): boolean;
export declare function expireAwarenessPeers(store: AwarenessPeerStore, now: number, timeout?: number): boolean;
export declare function awarenessPeers(store: AwarenessPeerStore): AwarenessPeer[];
export declare function resolveAwarenessCursor(cursor: AwarenessCursor, sheetIds: readonly string[], rows?: number, columns?: number): ResolvedAwarenessCursor | null;
/** @internal */
export declare function validAwarenessCursor(value: unknown, _maxFrameBytes?: number): value is AwarenessCursor;

// ---- collaboration/index.d.ts ----
export { CollaborationProvider } from './provider';
/** @internal */
export { attachCollaborationPipe, collaborationPipeFor } from './asyncReplica';
/** @experimental */
export { createAsyncCollaborationReplica } from './asyncReplica';
/** @experimental */
export type { AsyncCollaborationReplica } from './asyncReplica';
/** @experimental */
export type { AsyncReplicaAdapter } from './asyncReplica';
export { AWARENESS_BROADCAST_INTERVAL_MS, AWARENESS_COLORS, AWARENESS_HEARTBEAT_INTERVAL_MS, AWARENESS_LABEL_DURATION_MS, AWARENESS_PEER_TIMEOUT_MS, MAX_TRACKED_AWARENESS_PEERS, XLSX_MAX_COLUMNS, XLSX_MAX_ROWS, awarenessPeers, applyAwarenessUpdates, colorForClientId, expireAwarenessPeers, normalizeAwarenessColor, normalizeCollaborationUser, resolveAwarenessCursor, type ApplyAwarenessUpdateOptions, type AwarenessPeerEntry, type AwarenessPeerStore, type ResolvedAwarenessCursor, } from './awareness';
export { CollaborationError, type AwarenessCell, type AwarenessCursor, type AwarenessListener, type AwarenessPayload, type AwarenessPeer, type AwarenessState, type AwarenessUpdate, type CollaborationAwareness, type CollaborationErrorCode, type CollaborationErrorListener, type CollaborationProviderOptions, type CollaborationReplica, type CollaborationStatus, type CollaborationStatusChange, type CollaborationStatusListener, type CollaborationTransport, type CollaborationTransportEvent, type CollaborationUpdateOrigin, type CollaborationUser, type CollaborationUserOptions, } from './types';

// ---- collaboration/provider.d.ts ----
import type { PipeArrival } from '../../../../shared/collaboration-pipe/worker';
import { type AwarenessCursor, type AwarenessListener, type AwarenessPeer, type CollaborationErrorListener, type CollaborationProviderOptions, type CollaborationReplica, type CollaborationStatus, type CollaborationStatusListener, type CollaborationTransport } from './types';
/** @internal */
export declare function providerArrival(provider: CollaborationProvider): PipeArrival;
/** @internal */
export declare function providerLeave(provider: CollaborationProvider): {
    clock: number;
    frame: Uint8Array;
} | null;
export declare class CollaborationProvider {
    private readonly replica;
    private readonly piped;
    private readonly asyncReplica;
    private replicaTail;
    private replicaBytes;
    private updateBatch;
    private syncArrival;
    private openEpoch;
    private readonly transport;
    private readonly maxFrameBytes;
    private readonly maxMessagesPerFrame;
    private readonly maxPendingBytes;
    private readonly statusListeners;
    private readonly errorListeners;
    private readonly awarenessListeners;
    private readonly awarenessStore;
    private readonly awarenessUser;
    private readonly pendingFrames;
    private unsubscribeReplica;
    private unsubscribeTransport;
    private hasTransportSubscription;
    private connectionStatus;
    private isSynced;
    private wantsConnection;
    private isOpen;
    private isDestroyed;
    private epoch;
    private connectAttempt;
    private nextListenerId;
    private queuedBytes;
    private isFlushing;
    private transportCleanup;
    private statusRevision;
    private awarenessClock;
    private reportedClockExhaustion;
    private awarenessCursor;
    private lastAwarenessSentAt;
    private awarenessBroadcastTimer;
    private awarenessHeartbeatTimer;
    private awarenessExpiryTimer;
    private hasPublishedAwarenessForOpen;
    constructor(replica: CollaborationReplica, transport: CollaborationTransport, options?: CollaborationProviderOptions);
    get status(): CollaborationStatus;
    get synced(): boolean;
    get pendingBytes(): number;
    get peers(): readonly AwarenessPeer[];
    connect(): void;
    private connectPipedTransport;
    private startConnection;
    disconnect(): void;
    destroy(): void;
    onStatus(listener: CollaborationStatusListener): () => void;
    onError(listener: CollaborationErrorListener): () => void;
    onAwareness(listener: AwarenessListener): () => void;
    setCursor(cursor: AwarenessCursor | null): void;
    private isCurrent;
    private failConnection;
    private takeTransportSubscription;
    private cleanupTransport;
    private handleTransportEvent;
    private handleOpen;
    private sendHandshake;
    private sendFullUpdate;
    private handleMessage;
    private isCurrentOpen;
    private resetReplicaQueue;
    private enqueueReplicaMessages;
    /** Remote updates that arrive while the replica is busy join the queued batch and merge in one call. */
    private enqueueRemoteUpdates;
    private respondToSyncStep1;
    private beginArrival;
    private leaveFrame;
    private applyRemoteUpdate;
    private forwardLocalUpdate;
    private publishAwareness;
    private startAwarenessTimers;
    private stopAwarenessTimers;
    private clearAwarenessPeers;
    private emitAwareness;
    private sendFrame;
    private sendBestEffort;
    private queueFrame;
    private flushPending;
    private clearPending;
    private setStatus;
    private report;
}

// ---- collaboration/types.d.ts ----
export type CollaborationUpdateOrigin = 'local' | 'remote';
export interface CollaborationReplica {
    readonly clientId: number;
    encodeStateVector(): Uint8Array;
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
    applyUpdate(update: Uint8Array): unknown;
    onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
}
export type CollaborationTransportEvent = {
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
export interface CollaborationTransport {
    /** Starts one attempt; after close, reopen internally or require another provider.connect(). */
    connect(): void | Promise<void>;
    disconnect(): void | Promise<void>;
    /** Returns true only when the owned frame was accepted; false requires a later drain event. */
    send(data: Uint8Array): boolean;
    onEvent(listener: (event: CollaborationTransportEvent) => void): () => void;
}
export type CollaborationStatus = 'disconnected' | 'connecting' | 'connected' | 'destroyed';
export interface CollaborationStatusChange {
    status: CollaborationStatus;
    synced: boolean;
}
export interface CollaborationUser {
    name: string;
    color: string;
}
export interface CollaborationUserOptions {
    name: string;
    color?: string;
}
export interface AwarenessCell {
    row: number;
    col: number;
}
export interface AwarenessCursor {
    sheet: string;
    anchor: AwarenessCell;
    head: AwarenessCell;
}
export interface AwarenessPayload {
    user: CollaborationUser;
    cursor: AwarenessCursor | null;
}
export interface AwarenessState extends AwarenessPayload {
    clientId: number;
    clock: number;
}
export interface AwarenessUpdate {
    clientId: number;
    clock: number;
    state: AwarenessPayload | null;
}
export interface AwarenessPeer extends AwarenessState {
    lastSeen: number;
    cursorMovedAt: number;
}
export interface CollaborationAwareness {
    readonly peers: readonly AwarenessPeer[];
    setCursor(cursor: AwarenessCursor | null): void;
    onAwareness(listener: AwarenessListener): () => void;
}
export type CollaborationErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';
export declare class CollaborationError extends Error {
    readonly code: CollaborationErrorCode;
    readonly cause: unknown;
    constructor(code: CollaborationErrorCode, message: string, cause?: unknown);
}
export interface CollaborationProviderOptions {
    maxFrameBytes?: number;
    maxMessagesPerFrame?: number;
    maxPendingBytes?: number;
    user?: CollaborationUserOptions;
}
export type CollaborationStatusListener = (change: CollaborationStatusChange) => void;
export type CollaborationErrorListener = (error: CollaborationError) => void;
export type AwarenessListener = (peers: readonly AwarenessPeer[]) => void;

// ---- shared/collaboration-pipe/main.d.ts ----
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

// ---- shared/collaboration-pipe/protocol.d.ts ----
/**
 * Wire format between a main-thread collaboration facade and the provider its document worker runs.
 * Main owns the host transport's lifecycle and forwards its bytes untouched; the worker runs the protocol and
 * answers with frames to send, state and the transport calls its own failures need.
 */
export type PipeErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';
export interface PipeError {
    code: PipeErrorCode;
    message: string;
}
export interface PipeUser {
    name: string;
    color?: string;
}
export interface PipeProviderOptions {
    maxFrameBytes?: number;
    maxMessagesPerFrame?: number;
    syncChunkBytes?: number;
    maxSyncBytes?: number;
    user?: PipeUser;
}
/** The docx cursor; the default cursor type `C` of the pipe types. */
export interface PipeCursor {
    story: string;
    anchor: Uint8Array;
    head: Uint8Array;
}
export type PipeTransportEvent = {
    type: 'open';
} | {
    type: 'message';
    data: ArrayBuffer;
} | {
    type: 'close';
    reason?: string;
};
/**
 * Main to worker, in order; `seq` increases. `connect`, `disconnect`, `destroy` and `fail` replay what main already
 * did to the host transport, so the worker sends nothing and asks for no transport call while it runs them.
 */
export type PipeOp<C = PipeCursor> = {
    op: 'create';
    seq: number;
    options: PipeProviderOptions;
} | {
    op: 'connect';
    seq: number;
} | {
    op: 'disconnect';
    seq: number;
} | {
    op: 'destroy';
    seq: number;
} | {
    op: 'fail';
    seq: number;
    error: PipeError;
} | {
    op: 'cursor';
    seq: number;
    cursor: C | null;
    broadcast: boolean;
} | {
    op: 'event';
    seq: number;
    event: PipeTransportEvent;
};
/** The docx peer; the default peer type `P` of the pipe types. */
export interface PipePeer {
    clientId: number;
    clock: number;
    user: {
        name: string;
        color: string;
    };
    cursor: PipeCursor | null;
    lastSeenAt: number;
    cursorMovedAt: number;
    inferredCursor: {
        clientId: number;
        story: string;
        paraId: string;
        endOffset: number;
    } | null;
}
/**
 * Worker to main, in order. `through` is the last op the worker had run; main drops output older than its own
 * latest lifecycle call.
 */
export type PipeOutput<P = PipePeer> = {
    do: 'unsubscribe';
    through: number;
} | {
    do: 'disconnect';
    through: number;
} | {
    do: 'send';
    data: ArrayBuffer;
    through: number;
}
/** The awareness leave main sends when the host disconnects first; null while there is none to send. */
 | {
    do: 'leave';
    data: ArrayBuffer | null;
    through: number;
} | {
    do: 'status';
    status: 'disconnected' | 'connecting' | 'connected' | 'destroyed';
    synced: boolean;
    through: number;
} | {
    do: 'peers';
    peers: P[];
    timeOrigin: number;
    through: number;
} | {
    do: 'error';
    error: PipeError;
    through: number;
}
/** The worker has no provider under this id. */
 | {
    do: 'lost';
};
export interface PipeOutputBatch<P = PipePeer> {
    provider: number;
    out: PipeOutput<P>[];
}
export interface PipeCommandPayload<C = PipeCursor> {
    provider: number;
    ops: PipeOp<C>[];
}
export interface PipePayloadShape {
    /** Checks a non-null cursor; its byte fields stay within `maxFrameBytes`. */
    validCursor(cursor: unknown, maxFrameBytes: number): boolean;
    /** The keys a `create` op's options may carry. */
    optionKeys: ReadonlySet<string>;
}
/** Buffers a pipe payload holds, for transfer. */
export declare function pipeOpBuffers<C>(ops: readonly PipeOp<C>[]): ArrayBuffer[];
export declare function ownedBuffer(bytes: Uint8Array): ArrayBuffer;
export declare const PIPE_OPTION_KEYS: ReadonlySet<string>;
/** The docx cursor check. */
export declare function validPipeCursor(cursor: unknown, maxFrameBytes: number): boolean;
/** Shape check for a pipe command payload; frames past `maxFrameBytes` fail it. */
export declare function validPipePayload<C = PipeCursor>(payload: unknown, maxFrameBytes: number, shape: PipePayloadShape): payload is PipeCommandPayload<C>;
/** Epoch time of this context's provider clock zero: performance.now() when present, else Date.now(). */
export declare function clockOrigin(): number;

// ---- shared/collaboration-pipe/worker.d.ts ----
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

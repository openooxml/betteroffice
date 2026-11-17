// @betteroffice/docx declaration collaboration/provider.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type CollaborationCursor, type CollaborationErrorListener, type CollaborationPeer, type CollaborationPeerListener, type CollaborationProviderOptions, type CollaborationReplica, type CollaborationStatus, type CollaborationStatusListener, type CollaborationTransport } from './types';
import type { PipeArrival } from '../../../../shared/collaboration-pipe/worker';
/** @internal The awareness leave `provider` would send now, by the clock it carries; null when it would send none. */
export declare function providerLeave(provider: CollaborationProvider): {
    clock: number;
    frame: Uint8Array;
} | null;
/**
 * @internal Hands a merge the replica deferred to `provider`: settle with the merge's typing inference once it
 * merged, or fail it. A worker-hosted provider's replica defers a burst's merges to the burst's end.
 */
export declare function providerArrival(provider: CollaborationProvider): PipeArrival;
export declare class CollaborationProvider {
    private readonly replica;
    private readonly piped;
    private readonly asyncReplica;
    private replicaTail;
    private replicaBytes;
    private openEpoch;
    private readonly transport;
    private readonly maxFrameBytes;
    private readonly maxMessagesPerFrame;
    private readonly maxPendingBytes;
    private readonly syncChunkBytes;
    private readonly maxSyncBytes;
    private syncHeld;
    private syncHeldBytes;
    private syncJoinPending;
    private syncRequestOwed;
    private updateBatch;
    private readonly localUser;
    private readonly statusListeners;
    private readonly errorListeners;
    private readonly peerListeners;
    private readonly pendingFrames;
    private awarenessRecords;
    private readonly awarenessGenerations;
    private awarenessGeneration;
    private awarenessFloor;
    private readonly arrivalsInFlight;
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
    /** Whether the merge the replica is handed now may report typing; a SyncStep2 may not. */
    private arrivalInfers;
    private localClock;
    private reportedClockExhaustion;
    private localCursor;
    private lastAwarenessSentAt;
    private cursorBroadcastTimer;
    private heartbeatTimer;
    private expiryTimer;
    constructor(replica: CollaborationReplica, transport: CollaborationTransport, options?: CollaborationProviderOptions);
    get status(): CollaborationStatus;
    get synced(): boolean;
    get pendingBytes(): number;
    get peers(): readonly CollaborationPeer[];
    setCursor(cursor: CollaborationCursor | null, broadcast?: boolean): void;
    onPeers(listener: CollaborationPeerListener): () => void;
    connect(): void;
    disconnect(): void;
    destroy(): void;
    onStatus(listener: CollaborationStatusListener): () => void;
    onError(listener: CollaborationErrorListener): () => void;
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
    /** Queues a remote update; updates arriving while the replica is busy merge in one `applyUpdateChunks` call. */
    private enqueueRemoteUpdate;
    private respondToSyncStep1;
    /**
     * Whether the SyncStep1 owed to an opening peer waits for the held join to merge. Only a provider that chunks its
     * answers waits; one that does not asks at once with its current state vector, so a first joiner whose open nobody
     * answered still gets the SyncStep2 that ends its hold.
     */
    private defersSyncRequest;
    /**
     * Once the join merges, sends the SyncStep1 put off while it was held: the answer carries only what the join missed,
     * not a second, unchunked copy of the joined state.
     */
    private sendOwedSyncRequest;
    private sendWorkerSyncRequest;
    /**
     * A SyncStep1 sent before the held join merges asks for chunks: its state vector predates the join, so the answer can
     * be as large as the join.
     */
    private syncRequestVector;
    private sendSyncRequest;
    /**
     * Chunks only when `syncChunkBytes` is set and the SyncStep1 asks for it: that peer holds the chunks and applies them
     * as one change. A peer applying them one at a time, such as another peer in a room, would run its repairs on
     * partial states.
     */
    private chunksAnswer;
    /** Largest update a chunked sync answer may carry: `syncChunkBytes`, within one frame. */
    private syncChunkLimit;
    /**
     * Sends every update but the last as an Update and the last as the SyncStep2 that completes the answer, encoding one
     * frame at a time once all are known to fit.
     */
    private sendSyncAnswer;
    /** Holds a pre-sync update for the SyncStep2; past `maxSyncBytes` the connection fails. */
    private holdSyncUpdate;
    /** Ends holding for this connection and returns what was held. */
    private takeSyncHeld;
    private applySyncStep2;
    private applyRemoteUpdate;
    private beginArrival;
    private settleArrival;
    private setAwarenessRecords;
    private applyTypingInference;
    private forwardLocalUpdate;
    private nextAwarenessClock;
    private encodeLocalAwareness;
    private broadcastLocalAwareness;
    private scheduleCursorBroadcast;
    private clearCursorBroadcastTimer;
    private startPresence;
    private stopPresence;
    private leaveFrame;
    private sendLeaveBestEffort;
    private emitPeers;
    private sendFrame;
    private queueFrame;
    private flushPending;
    private clearPending;
    private setStatus;
    private report;
}

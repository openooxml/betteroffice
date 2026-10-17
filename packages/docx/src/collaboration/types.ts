export type CollaborationUpdateOrigin = 'local' | 'remote';

export interface CollaborationReplica {
  readonly clientId: number;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  /**
   * Optional. `encodeStateAsUpdate(remoteStateVector)` as ordered updates of at most `maxChunkBytes` each, except
   * one carrying a single larger struct; one update when the diff fits. Each integrates once the earlier ones have.
   */
  encodeStateAsUpdateChunks?(remoteStateVector: Uint8Array, maxChunkBytes: number): Uint8Array[];
  applyUpdate(update: Uint8Array): unknown;
  /** Optional. Applies one initial sync's ordered updates as a single change, as if they were one update. */
  applyUpdateChunks?(updates: readonly Uint8Array[]): unknown;
  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void;
}

/**
 * @experimental A replica whose state lives in a worker; reads and merges are asynchronous. Wrap it with
 * `createAsyncCollaborationReplica` before passing it to `CollaborationProvider`. The wrapper's synchronous
 * `encodeStateVector` and `encodeStateAsUpdate` throw; call these methods through `wrapper.async` instead.
 */
export interface AsyncCollaborationReplica {
  readonly clientId: number;
  encodeStateVector(): Promise<Uint8Array>;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Promise<Uint8Array>;
  /** Optional. See `CollaborationReplica.encodeStateAsUpdateChunks`. */
  encodeStateAsUpdateChunks?(remoteStateVector: Uint8Array, maxChunkBytes: number): Promise<Uint8Array[]>;
  /** Resolves once the worker has merged the update; the value is the typing inference, if any. */
  applyUpdate(update: Uint8Array): Promise<unknown>;
  /** Optional. Resolves once the worker has merged one initial sync's updates as a single change. */
  applyUpdateChunks?(updates: readonly Uint8Array[]): Promise<unknown>;
  onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
}

export type CollaborationTransportEvent =
  | { type: 'open' }
  | { type: 'message'; data: Uint8Array }
  | { type: 'close'; reason?: string }
  | { type: 'error'; error: unknown }
  | { type: 'drain' };

export interface CollaborationTransport {
  /** Starts one connection attempt; transports may also reopen while still subscribed. */
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

export type CollaborationErrorCode = 'backpressure' | 'protocol' | 'replica' | 'transport';

export class CollaborationError extends Error {
  readonly code: CollaborationErrorCode;
  readonly cause: unknown;

  constructor(code: CollaborationErrorCode, message: string, cause?: unknown) {
    super(message);
    this.name = 'CollaborationError';
    this.code = code;
    this.cause = cause;
  }
}

export interface CollaborationProviderOptions {
  /**
   * Largest frame sent or accepted. Default 16 MiB, the relay's limit; at most 64 MiB plus a 6-byte sync header, so
   * no frame carries an update over the 64 MiB per-update bound replicas enforce.
   */
  maxFrameBytes?: number;
  maxMessagesPerFrame?: number;
  /**
   * Bytes held while the transport refuses frames, and for an async replica received data waiting to merge, chunked
   * syncs included. Default 16 MiB, at most 256 MiB.
   */
  maxPendingBytes?: number;
  /**
   * Largest update in a chunked answer to SyncStep1; unset, every answer is one SyncStep2. A provider whose replica
   * implements `applyUpdateChunks` asks for chunks in its SyncStep1; with this set and a replica that implements
   * `encodeStateAsUpdateChunks`, the answer is Update messages closed by a SyncStep2. Set it only when the transport
   * delivers an answer to the asking peer alone (point-to-point). Room and broadcast transports, the relay included,
   * must leave it unset: other peers there would apply the chunks one at a time. At most `maxFrameBytes`. A struct
   * larger than a frame, such as an inline image, still fails to encode.
   */
  syncChunkBytes?: number;
  /**
   * Updates held before the first SyncStep2 when the replica implements `applyUpdateChunks`, which then applies them
   * with the SyncStep2 as one change; each held update also counts 64 bytes. Default 64 MiB, at most 256 MiB; more
   * fails the connection with `backpressure`. Relays that retain every update bound a room by their own capacity, so
   * large documents there join through `initialUpdate` instead.
   */
  maxSyncBytes?: number;
  user?: CollaborationUser;
}

export interface CollaborationUser {
  name: string;
  color?: string;
}

export interface CollaborationResolvedUser {
  name: string;
  color: string;
}

export interface CollaborationCursor {
  story: string;
  anchor: Uint8Array;
  head: Uint8Array;
}

export interface CollaborationTextInsertion {
  clientId: number;
  story: string;
  paraId: string;
  endOffset: number;
}

export interface CollaborationAwarenessState {
  clientId: number;
  clock: number;
  user: CollaborationResolvedUser;
  cursor: CollaborationCursor | null;
}

export interface CollaborationPeer extends CollaborationAwarenessState {
  lastSeenAt: number;
  cursorMovedAt: number;
  inferredCursor: CollaborationTextInsertion | null;
}

export type CollaborationPeerListener = (peers: readonly CollaborationPeer[]) => void;

export interface CollaborationPresence {
  readonly peers: readonly CollaborationPeer[];
  setCursor(cursor: CollaborationCursor | null, broadcast?: boolean): void;
  onPeers(listener: CollaborationPeerListener): () => void;
}

export type CollaborationStatusListener = (change: CollaborationStatusChange) => void;
export type CollaborationErrorListener = (error: CollaborationError) => void;

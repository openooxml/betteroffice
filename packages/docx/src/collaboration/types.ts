export type CollaborationUpdateOrigin = 'local' | 'remote';

export interface CollaborationReplica {
  readonly clientId: number;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): unknown;
  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void;
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
  maxFrameBytes?: number;
  maxMessagesPerFrame?: number;
  maxPendingBytes?: number;
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

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
  user?: CollaborationUserOptions;
}

export type CollaborationStatusListener = (change: CollaborationStatusChange) => void;
export type CollaborationErrorListener = (error: CollaborationError) => void;
export type AwarenessListener = (peers: readonly AwarenessPeer[]) => void;

import {
  decodeAwarenessUpdate,
  decodeMessages,
  DEFAULT_MAX_FRAME_BYTES,
  DEFAULT_MAX_MESSAGES_PER_FRAME,
  encodeAwarenessMessage,
  encodeQueryAwareness,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
  holdsOpenSyncClient,
  MAX_BUFFERED_BYTES,
  MAX_FRAME_BYTES,
  maxSyncPayloadBytes,
  syncFrameBytes,
  parseSyncRequest,
  withSyncRequest,
} from './protocol';
import {
  awarenessPeers,
  expireAwarenessRecords,
  MAX_AWARENESS_CURSOR_BYTES,
  MAX_AWARENESS_STRING_LENGTH,
  MAX_AWARENESS_TRACKED_PEERS,
  reduceAwarenessEntries,
  reduceTypingInference,
  resolvePresenceColor,
  type AwarenessRecord,
} from './awareness';
import {
  CollaborationError,
  type AsyncCollaborationReplica,
  type CollaborationCursor,
  type CollaborationErrorCode,
  type CollaborationErrorListener,
  type CollaborationPeer,
  type CollaborationPeerListener,
  type CollaborationProviderOptions,
  type CollaborationReplica,
  type CollaborationResolvedUser,
  type CollaborationStatus,
  type CollaborationStatusChange,
  type CollaborationStatusListener,
  type CollaborationTextInsertion,
  type CollaborationTransport,
  type CollaborationTransportEvent,
} from './types';
import { asyncReplicaFor } from './asyncReplica';

const DEFAULT_MAX_PENDING_BYTES = 16 * 1024 * 1024;
const DEFAULT_MAX_SYNC_BYTES = 64 * 1024 * 1024;
/** What a held update costs beyond its bytes, so empty updates cannot pile up for free. */
const HELD_UPDATE_OVERHEAD_BYTES = 64;
const CURSOR_BROADCAST_INTERVAL_MS = 80;
const HEARTBEAT_INTERVAL_MS = 20_000;
const PEER_EXPIRY_MS = 45_000;
const EXPIRY_SWEEP_INTERVAL_MS = 5_000;
type Timer = ReturnType<typeof setTimeout>;

function validateLimit(name: string, value: number, allowZero: boolean, max: number): number {
  if (!Number.isSafeInteger(value) || value < (allowZero ? 0 : 1) || value > max) {
    throw new RangeError(
      `${name} must be ${allowZero ? 'a non-negative' : 'a positive'} integer no greater than ${max}`
    );
  }
  return value;
}

function requireUpdates(value: unknown, operation: string): Uint8Array[] {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    !value.every((update) => update instanceof Uint8Array && update.byteLength > 0)
  ) {
    throw new TypeError(`${operation} must return a non-empty array of non-empty Uint8Array`);
  }
  return value;
}

function errorMessage(cause: unknown): string {
  if (cause instanceof Error && cause.message) return cause.message;
  if (typeof cause === 'string' && cause) return cause;
  return 'Unknown error';
}

function normalizeError(
  code: CollaborationErrorCode,
  context: string,
  cause: unknown
): CollaborationError {
  if (cause instanceof CollaborationError) return cause;
  return new CollaborationError(code, `${context}: ${errorMessage(cause)}`, cause);
}

function requireBytes(value: unknown, operation: string): Uint8Array {
  if (!(value instanceof Uint8Array)) {
    throw new TypeError(`${operation} must return a Uint8Array`);
  }
  return value.slice();
}

function nowMs(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function unrefTimer(timer: Timer): void {
  (timer as Timer & { unref?: () => void }).unref?.();
}

function resolveUser(
  user: CollaborationProviderOptions['user'],
  clientId: number
): CollaborationResolvedUser | null {
  if (!user) return null;
  if (!Number.isSafeInteger(clientId) || clientId < 0) {
    throw new RangeError('clientId must be a non-negative safe integer for awareness');
  }
  if (typeof user.name !== 'string' || user.name.trim().length === 0) {
    throw new TypeError('user.name must be a non-empty string');
  }
  return {
    name: user.name.trim().slice(0, MAX_AWARENESS_STRING_LENGTH),
    color: resolvePresenceColor(clientId, user.color),
  };
}

function ownedCursor(cursor: CollaborationCursor | null): CollaborationCursor | null {
  if (cursor === null) return null;
  if (
    typeof cursor.story !== 'string' ||
    !cursor.story ||
    !(cursor.anchor instanceof Uint8Array) ||
    !(cursor.head instanceof Uint8Array)
  ) {
    throw new TypeError('cursor requires a story and Uint8Array anchor/head');
  }
  return {
    story: cursor.story,
    anchor: cursor.anchor.slice(),
    head: cursor.head.slice(),
  };
}

function encodableCursor(cursor: CollaborationCursor): boolean {
  return (
    cursor.story.length <= MAX_AWARENESS_STRING_LENGTH &&
    cursor.anchor.byteLength <= MAX_AWARENESS_CURSOR_BYTES &&
    cursor.head.byteLength <= MAX_AWARENESS_CURSOR_BYTES
  );
}

function typingInference(value: unknown): CollaborationTextInsertion | null {
  if (typeof value !== 'object' || value === null) return null;
  const candidate = value as Partial<CollaborationTextInsertion>;
  if (
    !Number.isSafeInteger(candidate.clientId) ||
    (candidate.clientId ?? -1) < 0 ||
    typeof candidate.story !== 'string' ||
    typeof candidate.paraId !== 'string' ||
    !Number.isSafeInteger(candidate.endOffset) ||
    (candidate.endOffset ?? -1) < 0
  ) {
    return null;
  }
  return {
    clientId: candidate.clientId as number,
    story: candidate.story,
    paraId: candidate.paraId,
    endOffset: candidate.endOffset as number,
  };
}

type ReplicaFrame = { connection: number; bytes: number; remaining: number };
type UpdateBatch = {
  connection: number;
  bytes: number;
  entries: { update: Uint8Array; frame: ReplicaFrame; arrival: number }[];
};

function isReplicaMessage(message: ReturnType<typeof decodeMessages>[number]): boolean {
  return message.type === 'sync-step-1' || message.type === 'sync-step-2' || message.type === 'update';
}

/** Applies one sync's held updates and its SyncStep2 update; a replica without the method takes them one by one. */
function applySync(
  replica: Pick<CollaborationReplica, 'applyUpdate' | 'applyUpdateChunks'>,
  held: readonly Uint8Array[],
  update: Uint8Array
): unknown {
  if (held.length === 0) return replica.applyUpdate(update);
  return replica.applyUpdateChunks!([...held, update]);
}

export class CollaborationProvider {
  private readonly replica: CollaborationReplica;
  private readonly asyncReplica: AsyncCollaborationReplica | undefined;
  private replicaTail: Promise<void> = Promise.resolve();
  private replicaBytes = 0;
  private openEpoch = 0;
  private readonly transport: CollaborationTransport;
  private readonly maxFrameBytes: number;
  private readonly maxMessagesPerFrame: number;
  private readonly maxPendingBytes: number;
  private readonly syncChunkBytes: number | undefined;
  private readonly maxSyncBytes: number;
  private syncHeld: Uint8Array[] | null = null;
  private syncHeldBytes = 0;
  private syncJoinPending = false;
  private syncRequestOwed = false;
  private updateBatch: UpdateBatch | null = null;
  private readonly localUser: CollaborationResolvedUser | null;
  private readonly statusListeners = new Map<number, CollaborationStatusListener>();
  private readonly errorListeners = new Map<number, CollaborationErrorListener>();
  private readonly peerListeners = new Map<number, CollaborationPeerListener>();
  private readonly pendingFrames: Uint8Array[] = [];
  private awarenessRecords: ReadonlyMap<number, AwarenessRecord> = new Map();
  private readonly awarenessGenerations = new Map<number, number>();
  private awarenessGeneration = 0;
  private awarenessFloor = 0;
  private readonly arrivalsInFlight = new Map<number, number>();
  private unsubscribeReplica: () => void;
  private unsubscribeTransport: () => void = () => {};
  private hasTransportSubscription = false;
  private connectionStatus: CollaborationStatus = 'disconnected';
  private isSynced = false;
  private wantsConnection = false;
  private isOpen = false;
  private isDestroyed = false;
  private epoch = 0;
  private connectAttempt = 0;
  private nextListenerId = 0;
  private queuedBytes = 0;
  private isFlushing = false;
  private localClock = 0;
  private reportedClockExhaustion = false;
  private localCursor: CollaborationCursor | null = null;
  private lastAwarenessSentAt = Number.NEGATIVE_INFINITY;
  private cursorBroadcastTimer: Timer | null = null;
  private heartbeatTimer: Timer | null = null;
  private expiryTimer: Timer | null = null;

  constructor(
    replica: CollaborationReplica,
    transport: CollaborationTransport,
    options: CollaborationProviderOptions = {}
  ) {
    this.replica = replica;
    this.asyncReplica = asyncReplicaFor(replica);
    this.transport = transport;
    this.maxFrameBytes = validateLimit(
      'maxFrameBytes',
      options.maxFrameBytes ?? DEFAULT_MAX_FRAME_BYTES,
      false,
      MAX_FRAME_BYTES
    );
    this.maxMessagesPerFrame = validateLimit(
      'maxMessagesPerFrame',
      options.maxMessagesPerFrame ?? DEFAULT_MAX_MESSAGES_PER_FRAME,
      false,
      Number.MAX_SAFE_INTEGER
    );
    this.maxPendingBytes = validateLimit(
      'maxPendingBytes',
      options.maxPendingBytes ?? DEFAULT_MAX_PENDING_BYTES,
      true,
      MAX_BUFFERED_BYTES
    );
    this.syncChunkBytes = options.syncChunkBytes === undefined
      ? undefined
      : validateLimit('syncChunkBytes', options.syncChunkBytes, false, this.maxFrameBytes);
    this.maxSyncBytes = validateLimit(
      'maxSyncBytes',
      options.maxSyncBytes ?? DEFAULT_MAX_SYNC_BYTES,
      false,
      MAX_BUFFERED_BYTES
    );
    this.localUser = resolveUser(options.user, replica.clientId);

    try {
      this.unsubscribeReplica = replica.onUpdate((update, origin) => {
        if (origin === 'local') this.forwardLocalUpdate(update);
      });
    } catch (cause) {
      throw normalizeError('replica', 'Failed to observe replica updates', cause);
    }
  }

  get status(): CollaborationStatus {
    return this.connectionStatus;
  }

  get synced(): boolean {
    return this.isSynced;
  }

  get pendingBytes(): number {
    return this.queuedBytes;
  }

  get peers(): readonly CollaborationPeer[] {
    return awarenessPeers(this.awarenessRecords);
  }

  setCursor(cursor: CollaborationCursor | null, broadcast = true): void {
    if (this.isDestroyed || !this.localUser) return;
    const owned = ownedCursor(cursor);
    this.localCursor = !owned || encodableCursor(owned) ? owned : null;
    if (owned && !this.localCursor) {
      this.report(
        new CollaborationError(
          'protocol',
          `Presence cursor exceeds ${MAX_AWARENESS_STRING_LENGTH} story characters or ${MAX_AWARENESS_CURSOR_BYTES} position bytes; cursor dropped`
        )
      );
    }
    if (!broadcast) {
      this.clearCursorBroadcastTimer();
      return;
    }
    this.scheduleCursorBroadcast();
  }

  onPeers(listener: CollaborationPeerListener): () => void {
    if (this.isDestroyed) return () => {};
    if (typeof listener !== 'function') throw new TypeError('peer listener must be a function');
    const id = this.nextListenerId++;
    this.peerListeners.set(id, listener);
    listener(this.peers);
    return () => this.peerListeners.delete(id);
  }

  connect(): void {
    if (this.isDestroyed || this.isOpen || this.connectionStatus === 'connecting') return;

    this.wantsConnection = true;
    const token = this.hasTransportSubscription ? this.epoch : ++this.epoch;
    const attempt = ++this.connectAttempt;
    this.setStatus('connecting', false);
    if (!this.isCurrent(token)) return;

    if (!this.hasTransportSubscription) {
      let unsubscribe: () => void;
      try {
        unsubscribe = this.transport.onEvent((event) => this.handleTransportEvent(token, event));
        if (typeof unsubscribe !== 'function') {
          throw new TypeError('Transport event subscription must return a function');
        }
      } catch (cause) {
        this.failConnection(
          token,
          normalizeError('transport', 'Failed to subscribe to transport events', cause)
        );
        return;
      }

      if (!this.isCurrent(token)) {
        try {
          unsubscribe();
        } catch (cause) {
          if (!this.isDestroyed) {
            this.report(
              normalizeError('transport', 'Failed to unsubscribe from stale transport events', cause)
            );
          }
        }
        return;
      }
      this.unsubscribeTransport = unsubscribe;
      this.hasTransportSubscription = true;
    }

    try {
      const result = this.transport.connect();
      if (result) {
        void result.catch((cause) => {
          if (this.isCurrent(token) && this.connectAttempt === attempt) {
            this.failConnection(
              token,
              normalizeError('transport', 'Transport connect failed', cause)
            );
          }
        });
      }
    } catch (cause) {
      this.failConnection(token, normalizeError('transport', 'Transport connect failed', cause));
    }
  }

  disconnect(): void {
    if (this.isDestroyed) return;
    const active =
      this.wantsConnection ||
      this.isOpen ||
      this.hasTransportSubscription ||
      this.connectionStatus !== 'disconnected';
    if (!active) return;

    this.resetReplicaQueue();
    this.sendLeaveBestEffort();
    this.stopPresence();
    this.wantsConnection = false;
    this.isOpen = false;
    this.epoch += 1;
    this.connectAttempt += 1;
    const unsubscribe = this.takeTransportSubscription();
    this.clearPending();
    const cleanupErrors = this.cleanupTransport(unsubscribe, true);
    this.setStatus('disconnected', false);
    for (const error of cleanupErrors) this.report(error);
  }

  destroy(): void {
    if (this.isDestroyed) return;

    const active =
      this.wantsConnection ||
      this.isOpen ||
      this.hasTransportSubscription ||
      this.connectionStatus === 'connecting' ||
      this.connectionStatus === 'connected';
    this.resetReplicaQueue();
    this.sendLeaveBestEffort();
    this.stopPresence();
    this.isDestroyed = true;
    this.wantsConnection = false;
    this.isOpen = false;
    this.epoch += 1;
    this.connectAttempt += 1;
    const unsubscribe = this.takeTransportSubscription();
    this.clearPending();
    const cleanupErrors = this.cleanupTransport(unsubscribe, active);
    this.setStatus('destroyed', false);
    for (const error of cleanupErrors) this.report(error);
    try {
      this.unsubscribeReplica();
    } catch (cause) {
      this.report(normalizeError('replica', 'Failed to unsubscribe from replica updates', cause));
    }
    this.unsubscribeReplica = () => {};
    this.statusListeners.clear();
    this.errorListeners.clear();
    this.peerListeners.clear();
  }

  onStatus(listener: CollaborationStatusListener): () => void {
    if (this.isDestroyed) return () => {};
    if (typeof listener !== 'function') throw new TypeError('status listener must be a function');
    const id = this.nextListenerId++;
    this.statusListeners.set(id, listener);
    return () => this.statusListeners.delete(id);
  }

  onError(listener: CollaborationErrorListener): () => void {
    if (this.isDestroyed) return () => {};
    if (typeof listener !== 'function') throw new TypeError('error listener must be a function');
    const id = this.nextListenerId++;
    this.errorListeners.set(id, listener);
    return () => this.errorListeners.delete(id);
  }

  private isCurrent(token: number): boolean {
    return !this.isDestroyed && this.wantsConnection && token === this.epoch;
  }

  private failConnection(token: number, error: CollaborationError): void {
    if (!this.isCurrent(token)) return;

    this.resetReplicaQueue();
    this.stopPresence();
    this.wantsConnection = false;
    this.isOpen = false;
    this.epoch += 1;
    this.connectAttempt += 1;
    const unsubscribe = this.takeTransportSubscription();
    this.clearPending();
    const cleanupErrors = this.cleanupTransport(unsubscribe, true);
    this.setStatus('disconnected', false);
    this.report(error);
    for (const cleanupError of cleanupErrors) this.report(cleanupError);
  }

  private takeTransportSubscription(): () => void {
    const unsubscribe = this.unsubscribeTransport;
    this.unsubscribeTransport = () => {};
    this.hasTransportSubscription = false;
    return unsubscribe;
  }

  private cleanupTransport(unsubscribe: () => void, disconnect: boolean): CollaborationError[] {
    const errors: CollaborationError[] = [];
    try {
      unsubscribe();
    } catch (cause) {
      errors.push(
        normalizeError('transport', 'Failed to unsubscribe from transport events', cause)
      );
    }
    if (!disconnect) return errors;

    try {
      const result = this.transport.disconnect();
      if (result) {
        void result.catch((cause) => {
          if (!this.isDestroyed) {
            this.report(normalizeError('transport', 'Transport disconnect failed', cause));
          }
        });
      }
    } catch (cause) {
      errors.push(normalizeError('transport', 'Transport disconnect failed', cause));
    }
    return errors;
  }

  private handleTransportEvent(token: number, event: CollaborationTransportEvent): void {
    if (!this.isCurrent(token)) return;

    switch (event.type) {
      case 'open':
        this.handleOpen(token);
        break;
      case 'message':
        if (this.isOpen) this.handleMessage(token, event.data);
        break;
      case 'close':
        this.resetReplicaQueue();
        this.isOpen = false;
        this.connectAttempt += 1;
        this.stopPresence();
        this.clearPending();
        this.setStatus('disconnected', false);
        break;
      case 'error':
        this.failConnection(
          token,
          normalizeError('transport', 'Transport error', event.error)
        );
        break;
      case 'drain':
        if (this.isOpen) this.flushPending();
        break;
      default:
        this.failConnection(
          token,
          new CollaborationError('transport', 'Unknown transport event')
        );
    }
  }

  private handleOpen(token: number): void {
    if (this.isOpen) return;
    const connection = ++this.openEpoch;
    this.replicaTail = Promise.resolve();
    this.replicaBytes = 0;
    this.updateBatch = null;
    this.syncHeld =
      typeof (this.asyncReplica ?? this.replica).applyUpdateChunks === 'function' ? [] : null;
    this.syncHeldBytes = 0;
    this.syncJoinPending = false;
    this.syncRequestOwed = false;
    this.isOpen = true;
    this.clearPending();
    this.setStatus('connected', false);
    if (!this.isCurrent(token) || !this.isOpen) return;

    if (this.asyncReplica) {
      const opening = Promise.resolve().then(() => {
        if (!this.isCurrentOpen(token, connection)) return undefined;
        return this.asyncReplica!.encodeStateVector();
      });
      void opening.then((vector) => {
        if (!this.isCurrentOpen(token, connection)) return;
        this.sendHandshake(token, requireBytes(vector, 'encodeStateVector'), connection);
      }).catch((cause) => {
        if (this.isCurrentOpen(token, connection)) {
          this.failConnection(
            token,
            normalizeError('replica', 'Failed to encode replica state vector', cause)
          );
        }
      });
      return;
    }

    let stateVector: Uint8Array;
    try {
      stateVector = requireBytes(this.replica.encodeStateVector(), 'encodeStateVector');
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('replica', 'Failed to encode replica state vector', cause)
      );
      return;
    }

    this.sendHandshake(token, stateVector, connection);
  }

  private sendHandshake(token: number, stateVector: Uint8Array, connection: number): void {
    if (!this.isCurrentOpen(token, connection)) return;
    try {
      this.sendFrame(encodeSyncStep1(
        withSyncRequest(stateVector, { open: true, chunked: this.syncHeld !== null }),
        this.maxFrameBytes
      ));
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('protocol', 'Failed to encode collaboration handshake', cause)
      );
      return;
    }
    if (holdsOpenSyncClient(stateVector) && this.isCurrentOpen(token, connection)) {
      this.sendFullUpdate(token, connection);
    }
    if (!this.localUser || !this.isCurrentOpen(token, connection)) return;

    try {
      this.sendFrame(encodeQueryAwareness(this.maxFrameBytes));
    } catch (cause) {
      this.report(normalizeError('protocol', 'Failed to encode awareness query', cause));
    }
    if (!this.isCurrentOpen(token, connection)) return;
    this.broadcastLocalAwareness();
    if (!this.isCurrentOpen(token, connection)) return;
    this.startPresence();
  }

  private sendFullUpdate(token: number, connection: number): void {
    const send = (update: Uint8Array) => {
      if (!this.isCurrentOpen(token, connection)) return;
      try {
        this.sendFrame(encodeUpdate(update, this.maxFrameBytes));
      } catch (cause) {
        this.failConnection(token, normalizeError('protocol', 'Failed to encode collaboration handshake', cause));
      }
    };
    const fail = (cause: unknown) => {
      if (this.isCurrentOpen(token, connection)) {
        this.failConnection(token, normalizeError('replica', 'Failed to encode replica update', cause));
      }
    };
    if (this.asyncReplica) {
      void this.asyncReplica.encodeStateAsUpdate()
        .then((update) => requireBytes(update, 'encodeStateAsUpdate'))
        .then(send, fail);
      return;
    }
    let update: Uint8Array;
    try {
      update = requireBytes(this.replica.encodeStateAsUpdate(), 'encodeStateAsUpdate');
    } catch (cause) {
      fail(cause);
      return;
    }
    send(update);
  }

  private handleMessage(
    token: number,
    data: Uint8Array,
    decoded?: ReturnType<typeof decodeMessages>
  ): void {
    const connection = this.openEpoch;
    let messages: ReturnType<typeof decodeMessages>;
    try {
      messages = decoded ?? decodeMessages(data, this.maxFrameBytes, this.maxMessagesPerFrame);
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('protocol', 'Invalid collaboration frame', cause)
      );
      return;
    }

    if (this.asyncReplica && !decoded) {
      const frame = { connection, bytes: data.byteLength, remaining: 0 };
      let charged = false;
      for (const message of messages) {
        if (!this.isCurrentOpen(token, connection)) return;
        if (message.type === 'update' && this.syncHeld) {
          if (!this.holdSyncUpdate(token, message.update)) return;
        } else if (isReplicaMessage(message)) {
          if (!charged) {
            if (frame.bytes > this.maxPendingBytes - this.replicaBytes) {
              this.failConnection(
                token,
                new CollaborationError('backpressure', 'Pending replica data exceeds the byte limit')
              );
              return;
            }
            this.replicaBytes += frame.bytes;
            charged = true;
          }
          frame.remaining += 1;
          if (message.type === 'sync-step-2' && this.syncHeld) this.syncJoinPending = true;
          const held = message.type === 'sync-step-2' ? this.takeSyncHeld() : [];
          const owed = message.type === 'sync-step-1' && this.defersSyncRequest() &&
            parseSyncRequest(message.stateVector).open;
          if (owed) this.syncRequestOwed = true;
          if (message.type === 'update' && this.asyncReplica.applyUpdateChunks) {
            this.enqueueRemoteUpdate(token, frame, message.update);
          } else {
            this.enqueueReplicaMessages(token, frame, [message], held, owed);
          }
        } else {
          this.handleMessage(token, new Uint8Array(), [message]);
        }
      }
      return;
    }

    for (const message of messages) {
      if (!this.isCurrent(token) || !this.isOpen) return;

      switch (message.type) {
        case 'sync-step-1':
          this.respondToSyncStep1(token, connection, message.stateVector);
          break;
        case 'sync-step-2':
          if (this.applySyncStep2(token, message.update) && this.isCurrent(token)) {
            this.setStatus('connected', true);
            if (this.localUser) {
              try {
                this.sendFrame(encodeQueryAwareness(this.maxFrameBytes));
              } catch (cause) {
                this.report(
                  normalizeError('protocol', 'Failed to query awareness after sync', cause)
                );
              }
            }
            this.sendOwedSyncRequest(token, connection);
          }
          break;
        case 'update':
          if (this.syncHeld) this.holdSyncUpdate(token, message.update);
          else this.applyRemoteUpdate(token, message.update, true);
          break;
        case 'awareness': {
          const discarded: unknown[] = [];
          try {
            const reduction = reduceAwarenessEntries(
              this.awarenessRecords,
              decodeAwarenessUpdate(message.update, {
                trackedClientIds: this.awarenessRecords.keys(),
                localClientId: this.replica.clientId,
                onDiscard: (error) => discarded.push(error),
              }),
              this.replica.clientId,
              nowMs()
            );
            this.setAwarenessRecords(reduction.records);
            if (reduction.peersChanged) this.emitPeers();
          } catch (cause) {
            this.report(normalizeError('protocol', 'Failed to apply awareness update', cause));
          }
          for (const cause of discarded) {
            this.report(normalizeError('protocol', 'Discarded over-limit awareness data', cause));
          }
          break;
        }
        case 'auth':
          this.failConnection(
            token,
            new CollaborationError(
              'protocol',
              message.reason ? `Authentication denied: ${message.reason}` : 'Authentication denied'
            )
          );
          return;
        case 'query-awareness':
          try {
            if (this.localUser) this.broadcastLocalAwareness();
            else this.sendFrame(encodeAwarenessMessage([], this.maxFrameBytes));
          } catch (cause) {
            this.report(normalizeError('protocol', 'Failed to encode awareness response', cause));
          }
          break;
      }
    }
  }

  private isCurrentOpen(token: number, connection: number): boolean {
    return this.isCurrent(token) && this.isOpen && connection === this.openEpoch;
  }

  private resetReplicaQueue(): void {
    this.openEpoch += 1;
    this.replicaTail = Promise.resolve();
    this.replicaBytes = 0;
    this.updateBatch = null;
    this.syncHeld = null;
    this.syncHeldBytes = 0;
    this.syncJoinPending = false;
    this.syncRequestOwed = false;
    this.arrivalsInFlight.clear();
    this.awarenessGenerations.clear();
  }

  private enqueueReplicaMessages(
    token: number,
    frame: ReplicaFrame,
    messages: ReturnType<typeof decodeMessages>,
    held: readonly Uint8Array[] = [],
    owed = false
  ): void {
    const connection = frame.connection;
    if (!this.isCurrentOpen(token, connection)) return;
    this.updateBatch = null;
    const awarenessAtArrival = this.awarenessGeneration;
    this.arrivalsInFlight.set(awarenessAtArrival, (this.arrivalsInFlight.get(awarenessAtArrival) ?? 0) + 1);
    const operation = this.replicaTail.then(async () => {
      for (const message of messages) {
        if (!this.isCurrentOpen(token, connection)) return;
        try {
          if (message.type === 'sync-step-1') {
            const replica = this.asyncReplica!;
            const request = parseSyncRequest(message.stateVector);
            const updates = this.chunksAnswer(request.chunked, replica)
              ? requireUpdates(
                await replica.encodeStateAsUpdateChunks!(request.stateVector.slice(), this.syncChunkLimit()),
                'encodeStateAsUpdateChunks'
              )
              : [requireBytes(await replica.encodeStateAsUpdate(request.stateVector.slice()), 'encodeStateAsUpdate')];
            if (!this.isCurrentOpen(token, connection)) return;
            this.sendSyncAnswer(token, updates, () => this.isCurrentOpen(token, connection));
            if (request.open && !owed && this.isCurrentOpen(token, connection)) {
              await this.sendWorkerSyncRequest(token, connection);
            }
          } else if (message.type === 'sync-step-2' || message.type === 'update') {
            const applied = message.type === 'sync-step-2'
              ? await applySync(this.asyncReplica!, held, message.update)
              : await this.asyncReplica!.applyUpdate(message.update);
            if (!this.isCurrentOpen(token, connection)) return;
            if (message.type === 'sync-step-2') {
              this.syncJoinPending = false;
              this.setStatus('connected', true);
              if (!this.isCurrentOpen(token, connection)) return;
              if (this.localUser) {
                try {
                  this.sendFrame(encodeQueryAwareness(this.maxFrameBytes));
                } catch (cause) {
                  this.report(
                    normalizeError('protocol', 'Failed to query awareness after sync', cause)
                  );
                }
              }
              if (this.syncRequestOwed) {
                this.syncRequestOwed = false;
                await this.sendWorkerSyncRequest(token, connection);
              }
            } else {
              this.applyTypingInference(applied, awarenessAtArrival);
            }
          }
        } catch (cause) {
          if (this.isCurrentOpen(token, connection)) {
            this.failConnection(
              token,
              normalizeError(
                'replica',
                message.type === 'sync-step-1'
                  ? 'Failed to encode replica update'
                  : 'Failed to apply remote update',
                cause
              )
            );
          }
          return;
        }
      }
    }).finally(() => {
      frame.remaining -= messages.length;
      if (connection !== this.openEpoch) return;
      if (frame.remaining === 0) this.replicaBytes -= frame.bytes;
      this.settleArrival(awarenessAtArrival);
    });
    this.replicaTail = operation.catch(() => {});
  }

  /** Queues a remote update; updates arriving while the replica is busy merge in one `applyUpdateChunks` call. */
  private enqueueRemoteUpdate(token: number, frame: ReplicaFrame, update: Uint8Array): void {
    const connection = frame.connection;
    if (!this.isCurrentOpen(token, connection)) return;
    const arrival = this.awarenessGeneration;
    this.arrivalsInFlight.set(arrival, (this.arrivalsInFlight.get(arrival) ?? 0) + 1);
    const entry = { update, frame, arrival };
    const pending = this.updateBatch;
    if (pending?.connection === connection && pending.bytes + update.byteLength <= this.maxSyncBytes) {
      pending.entries.push(entry);
      pending.bytes += update.byteLength;
      return;
    }
    const batch: UpdateBatch = { connection, bytes: update.byteLength, entries: [entry] };
    this.updateBatch = batch;
    const operation = this.replicaTail.then(async () => {
      if (this.updateBatch === batch) this.updateBatch = null;
      if (!this.isCurrentOpen(token, connection)) return;
      const replica = this.asyncReplica!;
      try {
        if (batch.entries.length === 1) {
          const applied = await replica.applyUpdate(update);
          if (this.isCurrentOpen(token, connection)) this.applyTypingInference(applied, arrival);
        } else {
          await replica.applyUpdateChunks!(batch.entries.map((queued) => queued.update));
        }
      } catch (cause) {
        if (this.isCurrentOpen(token, connection)) {
          this.failConnection(token, normalizeError('replica', 'Failed to apply remote update', cause));
        }
      }
    }).finally(() => {
      for (const queued of batch.entries) {
        queued.frame.remaining -= 1;
        if (connection !== this.openEpoch) continue;
        if (queued.frame.remaining === 0) this.replicaBytes -= queued.frame.bytes;
        this.settleArrival(queued.arrival);
      }
    });
    this.replicaTail = operation.catch(() => {});
  }

  private respondToSyncStep1(token: number, connection: number, remoteStateVector: Uint8Array): void {
    const request = parseSyncRequest(remoteStateVector);
    let updates: Uint8Array[];
    try {
      updates = this.chunksAnswer(request.chunked, this.replica)
        ? requireUpdates(
          this.replica.encodeStateAsUpdateChunks!(request.stateVector.slice(), this.syncChunkLimit()),
          'encodeStateAsUpdateChunks'
        )
        : [requireBytes(this.replica.encodeStateAsUpdate(request.stateVector.slice()), 'encodeStateAsUpdate')];
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('replica', 'Failed to encode replica update', cause)
      );
      return;
    }
    this.sendSyncAnswer(token, updates, () => this.isCurrentOpen(token, connection));
    if (!request.open || !this.isCurrentOpen(token, connection)) return;
    if (this.defersSyncRequest()) this.syncRequestOwed = true;
    else this.sendSyncRequest(token);
  }

  /**
   * Whether the SyncStep1 owed to an opening peer waits for the held join to merge. Only a provider that chunks its
   * answers waits; one that does not asks at once with its current state vector, so a first joiner whose open nobody
   * answered still gets the SyncStep2 that ends its hold.
   */
  private defersSyncRequest(): boolean {
    return this.syncHeld !== null && this.syncChunkBytes !== undefined;
  }

  /**
   * Once the join merges, sends the SyncStep1 put off while it was held: the answer carries only what the join missed,
   * not a second, unchunked copy of the joined state.
   */
  private sendOwedSyncRequest(token: number, connection: number): void {
    if (!this.syncRequestOwed || !this.isCurrentOpen(token, connection)) return;
    this.syncRequestOwed = false;
    this.sendSyncRequest(token);
  }

  private async sendWorkerSyncRequest(token: number, connection: number): Promise<void> {
    const vector = requireBytes(await this.asyncReplica!.encodeStateVector(), 'encodeStateVector');
    if (this.isCurrentOpen(token, connection)) {
      this.sendFrame(encodeSyncStep1(this.syncRequestVector(vector), this.maxFrameBytes));
    }
  }

  /**
   * A SyncStep1 sent before the held join merges asks for chunks: its state vector predates the join, so the answer can
   * be as large as the join.
   */
  private syncRequestVector(stateVector: Uint8Array): Uint8Array {
    return this.syncHeld || this.syncJoinPending ? withSyncRequest(stateVector, { chunked: true }) : stateVector;
  }

  private sendSyncRequest(token: number): void {
    let stateVector: Uint8Array;
    try {
      stateVector = requireBytes(this.replica.encodeStateVector(), 'encodeStateVector');
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('replica', 'Failed to encode replica state vector', cause)
      );
      return;
    }
    try {
      this.sendFrame(encodeSyncStep1(this.syncRequestVector(stateVector), this.maxFrameBytes));
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('protocol', 'Failed to encode collaboration handshake', cause)
      );
    }
  }

  /**
   * Chunks only when `syncChunkBytes` is set and the SyncStep1 asks for it: that peer holds the chunks and applies them
   * as one change. A peer applying them one at a time, such as another peer in a room, would run its repairs on
   * partial states.
   */
  private chunksAnswer(
    requested: boolean,
    replica: Pick<CollaborationReplica | AsyncCollaborationReplica, 'encodeStateAsUpdateChunks'>
  ): boolean {
    return (
      this.syncChunkBytes !== undefined &&
      typeof replica.encodeStateAsUpdateChunks === 'function' &&
      requested
    );
  }

  /** Largest update a chunked sync answer may carry: `syncChunkBytes`, within one frame. */
  private syncChunkLimit(): number {
    return Math.max(1, Math.min(this.syncChunkBytes!, maxSyncPayloadBytes(this.maxFrameBytes)));
  }

  /**
   * Sends every update but the last as an Update and the last as the SyncStep2 that completes the answer, encoding one
   * frame at a time once all are known to fit.
   */
  private sendSyncAnswer(token: number, updates: readonly Uint8Array[], current: () => boolean): void {
    if (updates.some((update) => syncFrameBytes(update.byteLength) > this.maxFrameBytes)) {
      this.failConnection(
        token,
        new CollaborationError('protocol', `Failed to encode SyncStep2: Frame exceeds ${this.maxFrameBytes} bytes`)
      );
      return;
    }
    for (let index = 0; index < updates.length; index++) {
      if (!current()) return;
      let frame: Uint8Array;
      try {
        frame = index === updates.length - 1
          ? encodeSyncStep2(updates[index], this.maxFrameBytes)
          : encodeUpdate(updates[index], this.maxFrameBytes);
      } catch (cause) {
        this.failConnection(token, normalizeError('protocol', 'Failed to encode SyncStep2', cause));
        return;
      }
      this.sendFrame(frame);
    }
  }

  /** Holds a pre-sync update for the SyncStep2; past `maxSyncBytes` the connection fails. */
  private holdSyncUpdate(token: number, update: Uint8Array): boolean {
    const cost = update.byteLength + HELD_UPDATE_OVERHEAD_BYTES;
    if (cost > this.maxSyncBytes - this.syncHeldBytes) {
      this.failConnection(
        token,
        new CollaborationError('backpressure', `Initial sync exceeds ${this.maxSyncBytes} bytes`)
      );
      return false;
    }
    this.syncHeld!.push(update);
    this.syncHeldBytes += cost;
    return true;
  }

  /** Ends holding for this connection and returns what was held. */
  private takeSyncHeld(): Uint8Array[] {
    const held = this.syncHeld ?? [];
    this.syncHeld = null;
    this.syncHeldBytes = 0;
    return held;
  }

  private applySyncStep2(token: number, update: Uint8Array): boolean {
    const held = this.takeSyncHeld();
    try {
      applySync(this.replica, held, update.slice());
      return true;
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('replica', 'Failed to apply remote update', cause)
      );
      return false;
    }
  }

  private applyRemoteUpdate(token: number, update: Uint8Array, inferTyping: boolean): boolean {
    try {
      const applied = this.replica.applyUpdate(update.slice());
      if (inferTyping) this.applyTypingInference(applied);
      return true;
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('replica', 'Failed to apply remote update', cause)
      );
      return false;
    }
  }

  private settleArrival(generation: number): void {
    const remaining = this.arrivalsInFlight.get(generation)! - 1;
    if (remaining > 0) this.arrivalsInFlight.set(generation, remaining);
    else this.arrivalsInFlight.delete(generation);
    if (this.arrivalsInFlight.size === 0) {
      this.awarenessGenerations.clear();
      return;
    }
    const [oldest] = this.arrivalsInFlight.keys();
    for (const [clientId, changed] of this.awarenessGenerations) {
      if (changed <= oldest) this.awarenessGenerations.delete(clientId);
    }
  }

  private setAwarenessRecords(records: ReadonlyMap<number, AwarenessRecord>): void {
    if (this.arrivalsInFlight.size > 0 && records !== this.awarenessRecords) {
      for (const [clientId, record] of records) {
        if (record.clock === this.awarenessRecords.get(clientId)?.clock) continue;
        this.awarenessGenerations.set(clientId, ++this.awarenessGeneration);
      }
      for (const clientId of this.awarenessRecords.keys()) {
        if (records.has(clientId)) continue;
        this.awarenessGenerations.set(clientId, ++this.awarenessGeneration);
      }
      if (this.awarenessGenerations.size > MAX_AWARENESS_TRACKED_PEERS) {
        this.awarenessGenerations.clear();
        this.awarenessFloor = this.awarenessGeneration;
      }
    }
    this.awarenessRecords = records;
  }

  private applyTypingInference(applied: unknown, awarenessAtArrival?: number): void {
    const inference = typingInference(applied);
    if (!inference) return;
    if (
      awarenessAtArrival !== undefined &&
      (awarenessAtArrival < this.awarenessFloor ||
        (this.awarenessGenerations.get(inference.clientId) ?? 0) > awarenessAtArrival)
    ) return;
    const reduction = reduceTypingInference(
      this.awarenessRecords,
      inference,
      this.replica.clientId,
      nowMs()
    );
    this.awarenessRecords = reduction.records;
    if (reduction.peersChanged) this.emitPeers();
  }

  private forwardLocalUpdate(update: Uint8Array): void {
    if (this.isDestroyed || !this.isOpen || !this.wantsConnection) return;
    const token = this.epoch;

    let ownedUpdate: Uint8Array;
    try {
      ownedUpdate = requireBytes(update, 'Replica update');
    } catch (cause) {
      this.failConnection(token, normalizeError('replica', 'Invalid replica update', cause));
      return;
    }

    try {
      this.sendFrame(encodeUpdate(ownedUpdate, this.maxFrameBytes));
    } catch (cause) {
      this.failConnection(token, normalizeError('protocol', 'Failed to encode update', cause));
    }
  }

  private nextAwarenessClock(): number {
    if (this.localClock >= Number.MAX_SAFE_INTEGER) {
      throw new RangeError('Awareness clock exceeds Number.MAX_SAFE_INTEGER');
    }
    this.localClock += 1;
    return this.localClock;
  }

  private encodeLocalAwareness(
    user: CollaborationResolvedUser,
    clock: number,
    cursor: CollaborationCursor | null
  ): Uint8Array {
    return encodeAwarenessMessage(
      [{ clientId: this.replica.clientId, clock, state: { user, cursor } }],
      this.maxFrameBytes
    );
  }

  private broadcastLocalAwareness(): void {
    if (!this.localUser || !this.isOpen || this.isDestroyed || !this.wantsConnection) return;
    this.clearCursorBroadcastTimer();
    if (this.localClock >= Number.MAX_SAFE_INTEGER) {
      if (!this.reportedClockExhaustion) {
        this.reportedClockExhaustion = true;
        this.report(
          new CollaborationError('protocol', 'Awareness clock exceeds Number.MAX_SAFE_INTEGER')
        );
      }
      return;
    }

    const user = this.localUser;
    const clock = this.nextAwarenessClock();
    let frame: Uint8Array;
    try {
      frame = this.encodeLocalAwareness(user, clock, this.localCursor);
    } catch (cause) {
      const dropped = this.localCursor;
      this.localCursor = null;
      if (!dropped) {
        this.report(normalizeError('protocol', 'Failed to encode local awareness', cause));
        return;
      }
      try {
        frame = this.encodeLocalAwareness(user, clock, null);
      } catch (retryCause) {
        this.report(normalizeError('protocol', 'Failed to encode local awareness', retryCause));
        return;
      }
      this.report(normalizeError('protocol', 'Dropped an unencodable presence cursor', cause));
    }
    this.sendFrame(frame);
    this.lastAwarenessSentAt = nowMs();
  }

  private scheduleCursorBroadcast(): void {
    if (!this.localUser || !this.isOpen || this.isDestroyed || !this.wantsConnection) return;
    const remaining = CURSOR_BROADCAST_INTERVAL_MS - (nowMs() - this.lastAwarenessSentAt);
    if (remaining <= 0) {
      this.broadcastLocalAwareness();
      return;
    }
    if (this.cursorBroadcastTimer) return;
    this.cursorBroadcastTimer = setTimeout(() => {
      this.cursorBroadcastTimer = null;
      if (!this.isOpen || this.isDestroyed || !this.wantsConnection) return;
      this.broadcastLocalAwareness();
    }, remaining);
    unrefTimer(this.cursorBroadcastTimer);
  }

  private clearCursorBroadcastTimer(): void {
    if (this.cursorBroadcastTimer) clearTimeout(this.cursorBroadcastTimer);
    this.cursorBroadcastTimer = null;
  }

  private startPresence(): void {
    if (
      !this.localUser ||
      !this.isOpen ||
      this.isDestroyed ||
      !this.wantsConnection ||
      this.heartbeatTimer ||
      this.expiryTimer
    ) {
      return;
    }
    this.heartbeatTimer = setInterval(() => {
      this.broadcastLocalAwareness();
    }, HEARTBEAT_INTERVAL_MS);
    this.expiryTimer = setInterval(() => {
      const reduction = expireAwarenessRecords(
        this.awarenessRecords,
        nowMs(),
        PEER_EXPIRY_MS
      );
      this.setAwarenessRecords(reduction.records);
      if (reduction.peersChanged) this.emitPeers();
    }, EXPIRY_SWEEP_INTERVAL_MS);
    unrefTimer(this.heartbeatTimer);
    unrefTimer(this.expiryTimer);
  }

  private stopPresence(): void {
    this.clearCursorBroadcastTimer();
    if (this.heartbeatTimer) clearInterval(this.heartbeatTimer);
    if (this.expiryTimer) clearInterval(this.expiryTimer);
    this.heartbeatTimer = null;
    this.expiryTimer = null;
    const hadPeers = awarenessPeers(this.awarenessRecords).length > 0;
    this.setAwarenessRecords(new Map());
    this.awarenessGenerations.clear();
    if (hadPeers) this.emitPeers();
  }

  private sendLeaveBestEffort(): void {
    if (!this.localUser || !this.isOpen || this.isDestroyed || !this.wantsConnection) return;
    this.clearCursorBroadcastTimer();
    try {
      const frame = encodeAwarenessMessage(
        [
          {
            clientId: this.replica.clientId,
            clock: this.nextAwarenessClock(),
            state: null,
          },
        ],
        this.maxFrameBytes
      );
      this.transport.send(frame.slice());
    } catch (cause) {
      this.report(normalizeError('transport', 'Failed to send awareness leave', cause));
    }
  }

  private emitPeers(): void {
    const peers = this.peers;
    for (const [id, listener] of [...this.peerListeners]) {
      if (this.peerListeners.get(id) !== listener) continue;
      try {
        listener(peers);
      } catch {}
    }
  }

  private sendFrame(frame: Uint8Array): void {
    if (!this.isOpen || this.isDestroyed || !this.wantsConnection) return;
    const token = this.epoch;
    if (this.pendingFrames.length > 0) {
      this.queueFrame(token, frame);
      return;
    }

    let accepted: boolean;
    try {
      accepted = this.transport.send(frame.slice());
    } catch (cause) {
      this.failConnection(token, normalizeError('transport', 'Transport send failed', cause));
      return;
    }

    if (typeof accepted !== 'boolean') {
      this.failConnection(
        token,
        new CollaborationError('transport', 'Transport send must return a boolean')
      );
      return;
    }
    if (!accepted && this.isCurrent(token) && this.isOpen) this.queueFrame(token, frame);
  }

  private queueFrame(token: number, frame: Uint8Array): void {
    if (frame.byteLength > this.maxPendingBytes - this.queuedBytes) {
      this.failConnection(
        token,
        new CollaborationError(
          'backpressure',
          `Pending collaboration data exceeds ${this.maxPendingBytes} bytes`
        )
      );
      return;
    }

    const ownedFrame = frame.slice();
    this.pendingFrames.push(ownedFrame);
    this.queuedBytes += ownedFrame.byteLength;
  }

  private flushPending(): void {
    if (this.isFlushing) return;
    this.isFlushing = true;
    try {
      while (this.isOpen && !this.isDestroyed && this.pendingFrames.length > 0) {
        const token = this.epoch;
        const frame = this.pendingFrames[0];
        let accepted: boolean;
        try {
          accepted = this.transport.send(frame.slice());
        } catch (cause) {
          this.failConnection(token, normalizeError('transport', 'Transport send failed', cause));
          return;
        }

        if (typeof accepted !== 'boolean') {
          this.failConnection(
            token,
            new CollaborationError('transport', 'Transport send must return a boolean')
          );
          return;
        }
        if (!accepted) return;
        if (!this.isOpen || this.pendingFrames[0] !== frame) return;
        this.pendingFrames.shift();
        this.queuedBytes -= frame.byteLength;
      }
    } finally {
      this.isFlushing = false;
    }
  }

  private clearPending(): void {
    this.pendingFrames.length = 0;
    this.queuedBytes = 0;
  }

  private setStatus(status: CollaborationStatus, synced: boolean): void {
    if (this.connectionStatus === status && this.isSynced === synced) return;
    this.connectionStatus = status;
    this.isSynced = synced;
    const change: CollaborationStatusChange = { status, synced };
    for (const [id, listener] of [...this.statusListeners]) {
      if (this.statusListeners.get(id) !== listener) continue;
      try {
        listener(change);
      } catch {}
    }
  }

  private report(error: CollaborationError): void {
    for (const [id, listener] of [...this.errorListeners]) {
      if (this.errorListeners.get(id) !== listener) continue;
      try {
        listener(error);
      } catch {}
    }
  }
}

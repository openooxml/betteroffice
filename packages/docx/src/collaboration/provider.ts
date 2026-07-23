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
} from './protocol';
import {
  awarenessPeers,
  expireAwarenessRecords,
  reduceAwarenessEntries,
  reduceTypingInference,
  type AwarenessRecord,
} from './awareness';
import {
  CollaborationError,
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

const DEFAULT_MAX_PENDING_BYTES = 16 * 1024 * 1024;
const CURSOR_BROADCAST_INTERVAL_MS = 80;
const HEARTBEAT_INTERVAL_MS = 20_000;
const PEER_EXPIRY_MS = 45_000;
const EXPIRY_SWEEP_INTERVAL_MS = 5_000;
const PRESENCE_COLORS = [
  '#0B57D0',
  '#B3261E',
  '#137333',
  '#7B1FA2',
  '#A142F4',
  '#A04B00',
  '#006A6A',
  '#455A64',
] as const;

type Timer = ReturnType<typeof setTimeout>;

function validateLimit(name: string, value: number, allowZero: boolean): number {
  if (!Number.isSafeInteger(value) || value < (allowZero ? 0 : 1)) {
    throw new RangeError(`${name} must be ${allowZero ? 'a non-negative' : 'a positive'} integer`);
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
  if (user.color !== undefined && (typeof user.color !== 'string' || !user.color.trim())) {
    throw new TypeError('user.color must be a non-empty string');
  }
  return {
    name: user.name.trim(),
    color: user.color?.trim() ?? PRESENCE_COLORS[clientId % PRESENCE_COLORS.length],
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

export class CollaborationProvider {
  private readonly replica: CollaborationReplica;
  private readonly transport: CollaborationTransport;
  private readonly maxFrameBytes: number;
  private readonly maxMessagesPerFrame: number;
  private readonly maxPendingBytes: number;
  private readonly localUser: CollaborationResolvedUser | null;
  private readonly statusListeners = new Map<number, CollaborationStatusListener>();
  private readonly errorListeners = new Map<number, CollaborationErrorListener>();
  private readonly peerListeners = new Map<number, CollaborationPeerListener>();
  private readonly pendingFrames: Uint8Array[] = [];
  private awarenessRecords: ReadonlyMap<number, AwarenessRecord> = new Map();
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
    this.transport = transport;
    this.maxFrameBytes = validateLimit(
      'maxFrameBytes',
      options.maxFrameBytes ?? DEFAULT_MAX_FRAME_BYTES,
      false
    );
    this.maxMessagesPerFrame = validateLimit(
      'maxMessagesPerFrame',
      options.maxMessagesPerFrame ?? DEFAULT_MAX_MESSAGES_PER_FRAME,
      false
    );
    this.maxPendingBytes = validateLimit(
      'maxPendingBytes',
      options.maxPendingBytes ?? DEFAULT_MAX_PENDING_BYTES,
      true
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
    this.localCursor = ownedCursor(cursor);
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
    this.isOpen = true;
    this.clearPending();
    this.setStatus('connected', false);
    if (!this.isCurrent(token) || !this.isOpen) return;

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
      this.sendFrame(encodeSyncStep1(stateVector, this.maxFrameBytes));
      if (this.localUser && this.isCurrent(token) && this.isOpen) {
        this.sendFrame(encodeQueryAwareness(this.maxFrameBytes));
        this.broadcastLocalAwareness();
        this.startPresence();
      }
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('protocol', 'Failed to encode collaboration handshake', cause)
      );
    }
  }

  private handleMessage(token: number, data: Uint8Array): void {
    let messages: ReturnType<typeof decodeMessages>;
    try {
      messages = decodeMessages(data.slice(), this.maxFrameBytes, this.maxMessagesPerFrame);
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('protocol', 'Invalid collaboration frame', cause)
      );
      return;
    }

    for (const message of messages) {
      if (!this.isCurrent(token) || !this.isOpen) return;

      switch (message.type) {
        case 'sync-step-1':
          this.respondToSyncStep1(token, message.stateVector);
          break;
        case 'sync-step-2':
          if (this.applyRemoteUpdate(token, message.update, false) && this.isCurrent(token)) {
            this.setStatus('connected', true);
            if (this.localUser) {
              try {
                this.sendFrame(encodeQueryAwareness(this.maxFrameBytes));
              } catch (cause) {
                this.failConnection(
                  token,
                  normalizeError('protocol', 'Failed to query awareness after sync', cause)
                );
                return;
              }
            }
          }
          break;
        case 'update':
          this.applyRemoteUpdate(token, message.update, true);
          break;
        case 'awareness':
          try {
            const reduction = reduceAwarenessEntries(
              this.awarenessRecords,
              decodeAwarenessUpdate(message.update),
              this.replica.clientId,
              nowMs()
            );
            this.awarenessRecords = reduction.records;
            if (reduction.peersChanged) this.emitPeers();
          } catch (cause) {
            this.failConnection(
              token,
              normalizeError('protocol', 'Failed to apply awareness update', cause)
            );
            return;
          }
          break;
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
            this.failConnection(
              token,
              normalizeError('protocol', 'Failed to encode awareness response', cause)
            );
            return;
          }
          break;
      }
    }
  }

  private respondToSyncStep1(token: number, remoteStateVector: Uint8Array): void {
    let update: Uint8Array;
    try {
      update = requireBytes(
        this.replica.encodeStateAsUpdate(remoteStateVector.slice()),
        'encodeStateAsUpdate'
      );
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('replica', 'Failed to encode replica update', cause)
      );
      return;
    }

    try {
      this.sendFrame(encodeSyncStep2(update, this.maxFrameBytes));
    } catch (cause) {
      this.failConnection(token, normalizeError('protocol', 'Failed to encode SyncStep2', cause));
    }
  }

  private applyRemoteUpdate(token: number, update: Uint8Array, inferTyping: boolean): boolean {
    try {
      const applied = this.replica.applyUpdate(update.slice());
      const inference = inferTyping ? typingInference(applied) : null;
      if (inference) {
        const reduction = reduceTypingInference(
          this.awarenessRecords,
          inference,
          this.replica.clientId,
          nowMs()
        );
        this.awarenessRecords = reduction.records;
        if (reduction.peersChanged) this.emitPeers();
      }
      return true;
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('replica', 'Failed to apply remote update', cause)
      );
      return false;
    }
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

  private broadcastLocalAwareness(): void {
    if (!this.localUser || !this.isOpen || this.isDestroyed || !this.wantsConnection) return;
    this.clearCursorBroadcastTimer();
    const clock = this.nextAwarenessClock();
    this.sendFrame(
      encodeAwarenessMessage(
        [
          {
            clientId: this.replica.clientId,
            clock,
            state: {
              user: this.localUser,
              cursor: this.localCursor,
            },
          },
        ],
        this.maxFrameBytes
      )
    );
    this.lastAwarenessSentAt = nowMs();
  }

  private scheduleCursorBroadcast(): void {
    if (!this.localUser || !this.isOpen || this.isDestroyed || !this.wantsConnection) return;
    const remaining = CURSOR_BROADCAST_INTERVAL_MS - (nowMs() - this.lastAwarenessSentAt);
    if (remaining <= 0) {
      try {
        this.broadcastLocalAwareness();
      } catch (cause) {
        this.failConnection(
          this.epoch,
          normalizeError('protocol', 'Failed to encode cursor awareness', cause)
        );
      }
      return;
    }
    if (this.cursorBroadcastTimer) return;
    this.cursorBroadcastTimer = setTimeout(() => {
      this.cursorBroadcastTimer = null;
      if (!this.isOpen || this.isDestroyed || !this.wantsConnection) return;
      try {
        this.broadcastLocalAwareness();
      } catch (cause) {
        this.failConnection(
          this.epoch,
          normalizeError('protocol', 'Failed to encode cursor awareness', cause)
        );
      }
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
      try {
        this.broadcastLocalAwareness();
      } catch (cause) {
        this.failConnection(
          this.epoch,
          normalizeError('protocol', 'Failed to encode awareness heartbeat', cause)
        );
      }
    }, HEARTBEAT_INTERVAL_MS);
    this.expiryTimer = setInterval(() => {
      const reduction = expireAwarenessRecords(
        this.awarenessRecords,
        nowMs(),
        PEER_EXPIRY_MS
      );
      this.awarenessRecords = reduction.records;
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
    this.awarenessRecords = new Map();
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

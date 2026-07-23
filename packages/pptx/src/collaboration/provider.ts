import {
  decodeAwarenessUpdate,
  decodeMessages,
  DEFAULT_MAX_FRAME_BYTES,
  DEFAULT_MAX_MESSAGES_PER_FRAME,
  encodeAwarenessUpdate,
  encodeQueryAwareness,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
} from './protocol';
import {
  PRESENCE_CURSOR_INTERVAL_MS,
  PRESENCE_EXPIRY_MS,
  PRESENCE_HEARTBEAT_MS,
  PresencePeers,
  presenceUser,
  samePresenceCursor,
} from './presence';
import {
  CollaborationError,
  type CollaborationErrorCode,
  type CollaborationErrorListener,
  type CollaborationProviderOptions,
  type CollaborationReplica,
  type CollaborationStatus,
  type CollaborationStatusChange,
  type CollaborationStatusListener,
  type CollaborationTransport,
  type CollaborationTransportEvent,
  type PptxPresenceCursor,
  type PptxPresenceListener,
  type PptxPresencePeer,
  type PptxPresenceState,
  type PptxPresenceUser,
} from './types';

const DEFAULT_MAX_PENDING_BYTES = 16 * 1024 * 1024;

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

function unrefTimer(timer: ReturnType<typeof setTimeout>): void {
  if (typeof timer === 'object' && timer && 'unref' in timer) {
    (timer as { unref(): void }).unref();
  }
}

function requireClientId(value: number): number {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError('Replica clientId must be a non-negative safe integer');
  }
  return value;
}

function normalizeCursor(cursor: PptxPresenceCursor | null): PptxPresenceCursor | null {
  if (cursor === null) return null;
  if (!cursor || typeof cursor.slideId !== 'string' || cursor.slideId.length === 0) {
    throw new TypeError('Presence cursor slideId must be a non-empty string');
  }
  if (
    cursor.shapeId !== undefined &&
    (typeof cursor.shapeId !== 'string' || cursor.shapeId.length === 0)
  ) {
    throw new TypeError('Presence cursor shapeId must be a non-empty string');
  }
  return cursor.shapeId
    ? { slideId: cursor.slideId, shapeId: cursor.shapeId }
    : { slideId: cursor.slideId };
}

export class CollaborationProvider {
  private readonly replica: CollaborationReplica;
  private readonly transport: CollaborationTransport;
  private readonly maxFrameBytes: number;
  private readonly maxMessagesPerFrame: number;
  private readonly maxPendingBytes: number;
  private readonly localClientId: number;
  private readonly localUser: PptxPresenceUser;
  private readonly remotePresence: PresencePeers;
  private readonly statusListeners = new Map<number, CollaborationStatusListener>();
  private readonly errorListeners = new Map<number, CollaborationErrorListener>();
  private readonly presenceListeners = new Map<number, PptxPresenceListener>();
  private readonly pendingFrames: Uint8Array[] = [];
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
  private localCursor: PptxPresenceCursor | null = null;
  private lastAwarenessSentAt = 0;
  private cursorTimer: ReturnType<typeof setTimeout> | null = null;
  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;
  private expiryTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(
    replica: CollaborationReplica,
    transport: CollaborationTransport,
    options: CollaborationProviderOptions = {}
  ) {
    this.replica = replica;
    this.transport = transport;
    this.localClientId = requireClientId(replica.clientId);
    this.localUser = presenceUser(this.localClientId, options.user);
    this.remotePresence = new PresencePeers(this.localClientId);
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

  get peers(): readonly PptxPresencePeer[] {
    return this.remotePresence.peers;
  }

  setCursor(cursor: PptxPresenceCursor | null): void {
    if (this.isDestroyed) return;
    const next = normalizeCursor(cursor);
    if (samePresenceCursor(this.localCursor, next)) return;
    this.localCursor = next;
    if (this.isOpen && this.wantsConnection) this.scheduleCursorBroadcast();
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

    if (this.isOpen && this.wantsConnection) {
      const token = this.epoch;
      this.broadcastLocalPresence(true);
      if (!this.isCurrent(token)) return;
    }
    this.stopPresenceTimers();
    this.wantsConnection = false;
    this.isOpen = false;
    this.epoch += 1;
    this.connectAttempt += 1;
    const unsubscribe = this.takeTransportSubscription();
    this.clearPending();
    this.clearRemotePresence();
    const cleanupErrors = this.cleanupTransport(unsubscribe, true);
    this.setStatus('disconnected', false);
    for (const error of cleanupErrors) this.report(error);
  }

  destroy(): void {
    if (this.isDestroyed) return;

    if (this.isOpen && this.wantsConnection) this.broadcastLocalPresence(true);
    if (this.isDestroyed) return;
    const active =
      this.wantsConnection ||
      this.isOpen ||
      this.hasTransportSubscription ||
      this.connectionStatus === 'connecting' ||
      this.connectionStatus === 'connected';
    this.stopPresenceTimers();
    this.isDestroyed = true;
    this.wantsConnection = false;
    this.isOpen = false;
    this.epoch += 1;
    this.connectAttempt += 1;
    const unsubscribe = this.takeTransportSubscription();
    this.clearPending();
    this.clearRemotePresence();
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
    this.presenceListeners.clear();
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

  onPresence(listener: PptxPresenceListener): () => void {
    if (this.isDestroyed) return () => {};
    if (typeof listener !== 'function') throw new TypeError('presence listener must be a function');
    const id = this.nextListenerId++;
    this.presenceListeners.set(id, listener);
    try {
      listener(this.peers);
    } catch {}
    return () => this.presenceListeners.delete(id);
  }

  private isCurrent(token: number): boolean {
    return !this.isDestroyed && this.wantsConnection && token === this.epoch;
  }

  private failConnection(token: number, error: CollaborationError): void {
    if (!this.isCurrent(token)) return;

    this.wantsConnection = false;
    this.isOpen = false;
    this.epoch += 1;
    this.connectAttempt += 1;
    this.stopPresenceTimers();
    const unsubscribe = this.takeTransportSubscription();
    this.clearPending();
    this.clearRemotePresence();
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
        this.stopPresenceTimers();
        this.clearPending();
        this.clearRemotePresence();
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
    } catch (cause) {
      this.failConnection(token, normalizeError('protocol', 'Failed to encode SyncStep1', cause));
      return;
    }
    if (!this.isCurrent(token) || !this.isOpen) return;
    this.startPresenceTimers();
    this.broadcastLocalPresence(false);
    if (!this.isCurrent(token) || !this.isOpen) return;
    try {
      this.sendFrame(encodeQueryAwareness(this.maxFrameBytes));
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('protocol', 'Failed to encode awareness query', cause)
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
          if (this.applyRemoteUpdate(token, message.update) && this.isCurrent(token)) {
            this.setStatus('connected', true);
          }
          break;
        case 'update':
          this.applyRemoteUpdate(token, message.update);
          break;
        case 'awareness': {
          try {
            const entries = decodeAwarenessUpdate(message.update, this.maxMessagesPerFrame);
            if (this.remotePresence.apply(entries, Date.now())) {
              this.emitPresence();
              this.schedulePresenceExpiry();
            }
          } catch (cause) {
            this.report(normalizeError('protocol', 'Invalid awareness update', cause));
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
          this.broadcastLocalPresence(false);
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

  private applyRemoteUpdate(token: number, update: Uint8Array): boolean {
    try {
      this.replica.applyUpdate(update.slice());
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

  private scheduleCursorBroadcast(): void {
    if (this.cursorTimer) return;
    const delay = Math.max(
      0,
      PRESENCE_CURSOR_INTERVAL_MS - (Date.now() - this.lastAwarenessSentAt)
    );
    if (delay === 0) {
      this.broadcastLocalPresence(false);
      return;
    }
    this.cursorTimer = setTimeout(() => {
      this.cursorTimer = null;
      this.broadcastLocalPresence(false);
    }, delay);
    unrefTimer(this.cursorTimer);
  }

  private broadcastLocalPresence(leaving: boolean): void {
    if (!this.isOpen || this.isDestroyed || !this.wantsConnection) return;
    const token = this.epoch;
    if (this.localClock >= Number.MAX_SAFE_INTEGER) {
      this.failConnection(
        token,
        new CollaborationError('protocol', 'Awareness clock exceeds Number.MAX_SAFE_INTEGER')
      );
      return;
    }
    this.localClock += 1;
    const clock = this.localClock;
    const state: PptxPresenceState | null = leaving
      ? null
      : {
          clientId: this.localClientId,
          clock,
          user: { ...this.localUser },
          cursor: this.localCursor ? { ...this.localCursor } : null,
        };
    if (this.cursorTimer) clearTimeout(this.cursorTimer);
    this.cursorTimer = null;
    this.lastAwarenessSentAt = Date.now();
    try {
      const frame = encodeAwarenessUpdate(
        [{ clientId: this.localClientId, clock, state }],
        this.maxFrameBytes
      );
      if (leaving) this.sendLeaveFrameBestEffort(frame);
      else this.sendFrame(frame);
    } catch (cause) {
      this.failConnection(
        token,
        normalizeError('protocol', 'Failed to encode awareness update', cause)
      );
    }
  }

  private sendLeaveFrameBestEffort(frame: Uint8Array): void {
    try {
      this.transport.send(frame.slice());
    } catch {}
  }

  private startPresenceTimers(): void {
    this.stopPresenceTimers();
    this.heartbeatTimer = setInterval(() => {
      this.broadcastLocalPresence(false);
    }, PRESENCE_HEARTBEAT_MS);
    unrefTimer(this.heartbeatTimer);
    this.schedulePresenceExpiry();
  }

  private stopPresenceTimers(): void {
    if (this.cursorTimer) clearTimeout(this.cursorTimer);
    if (this.heartbeatTimer) clearInterval(this.heartbeatTimer);
    if (this.expiryTimer) clearTimeout(this.expiryTimer);
    this.cursorTimer = null;
    this.heartbeatTimer = null;
    this.expiryTimer = null;
  }

  private schedulePresenceExpiry(): void {
    if (this.expiryTimer) clearTimeout(this.expiryTimer);
    this.expiryTimer = null;
    if (!this.isOpen || this.isDestroyed) return;
    const peers = this.remotePresence.peers;
    if (peers.length === 0) return;
    const expiresAt = Math.min(...peers.map((peer) => peer.lastSeen + PRESENCE_EXPIRY_MS));
    this.expiryTimer = setTimeout(() => {
      this.expiryTimer = null;
      if (this.remotePresence.expire(Date.now())) this.emitPresence();
      this.schedulePresenceExpiry();
    }, Math.max(0, expiresAt - Date.now()));
    unrefTimer(this.expiryTimer);
  }

  private clearRemotePresence(): void {
    if (this.remotePresence.clear()) this.emitPresence();
  }

  private emitPresence(): void {
    const peers = this.peers;
    for (const [id, listener] of [...this.presenceListeners]) {
      if (this.presenceListeners.get(id) !== listener) continue;
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

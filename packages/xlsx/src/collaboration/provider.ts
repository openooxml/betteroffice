import {
  decodeMessages,
  DEFAULT_MAX_FRAME_BYTES,
  DEFAULT_MAX_MESSAGES_PER_FRAME,
  encodeEmptyAwarenessUpdate,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
} from './protocol';
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
} from './types';

const DEFAULT_MAX_PENDING_BYTES = DEFAULT_MAX_FRAME_BYTES;

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

export class CollaborationProvider {
  private readonly replica: CollaborationReplica;
  private readonly transport: CollaborationTransport;
  private readonly maxFrameBytes: number;
  private readonly maxMessagesPerFrame: number;
  private readonly maxPendingBytes: number;
  private readonly statusListeners = new Map<number, CollaborationStatusListener>();
  private readonly errorListeners = new Map<number, CollaborationErrorListener>();
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
  private transportCleanup: Promise<void> | undefined;
  private statusRevision = 0;

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

  connect(): void {
    if (this.isDestroyed || this.isOpen || this.connectionStatus === 'connecting') return;

    this.wantsConnection = true;
    this.setStatus('connecting', false);
    if (this.transportCleanup || this.isDestroyed || !this.wantsConnection) return;
    this.startConnection();
  }

  private startConnection(): void {
    if (this.isDestroyed || !this.wantsConnection || this.isOpen || this.transportCleanup) return;
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
    if (this.transportCleanup) return errors;

    try {
      const result = this.transport.disconnect();
      if (result) {
        let failed = false;
        const cleanup = Promise.resolve(result)
          .catch((cause) => {
            failed = true;
            if (!this.isDestroyed) {
              this.report(normalizeError('transport', 'Transport disconnect failed', cause));
            }
          })
          .finally(() => {
            if (this.transportCleanup !== cleanup) return;
            this.transportCleanup = undefined;
            if (this.isDestroyed) return;
            if (failed) {
              this.wantsConnection = false;
              this.setStatus('disconnected', false);
            } else if (this.wantsConnection) {
              this.startConnection();
            }
          });
        this.transportCleanup = cleanup;
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
    } catch (cause) {
      this.failConnection(token, normalizeError('protocol', 'Failed to encode SyncStep1', cause));
    }
  }

  private handleMessage(token: number, data: Uint8Array): void {
    let messages: ReturnType<typeof decodeMessages>;
    try {
      if (!(data instanceof Uint8Array)) throw new TypeError('Frame must be a Uint8Array');
      if (data.byteLength > this.maxFrameBytes) {
        throw new RangeError(`Frame exceeds ${this.maxFrameBytes} bytes`);
      }
      messages = decodeMessages(data, this.maxFrameBytes, this.maxMessagesPerFrame);
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
        case 'awareness':
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
            this.sendFrame(encodeEmptyAwarenessUpdate(this.maxFrameBytes));
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
    const revision = ++this.statusRevision;
    const change: CollaborationStatusChange = { status, synced };
    for (const [id, listener] of [...this.statusListeners]) {
      if (this.statusRevision !== revision) return;
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

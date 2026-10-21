import {
  pipeOpBuffers,
  type PipeCursor,
  type PipeError,
  type PipeErrorCode,
  type PipeOp,
  type PipeOutput,
  type PipePeer,
  type PipeProviderOptions,
} from './protocol';

type Status = 'disconnected' | 'connecting' | 'connected' | 'destroyed';
type TransportEvent =
  | { type: 'open' }
  | { type: 'message'; data: Uint8Array }
  | { type: 'close'; reason?: string }
  | { type: 'error'; error: unknown }
  | { type: 'drain' };

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

export interface PipedCollaborationOptions<E> {
  worker: PipeProviderOptions;
  maxPendingBytes: number;
  hasUser: boolean;
  makeError(code: PipeErrorCode, message: string, cause?: unknown): E;
  /** `cause` itself when it already is an `E`, else a new `E` with `context` and the cause's message. */
  normalize(code: PipeErrorCode, context: string, cause: unknown): E;
}

function message(cause: unknown): string {
  if (cause instanceof Error && cause.message) return cause.message;
  if (typeof cause === 'string' && cause) return cause;
  return 'Unknown error';
}

const ERROR_CODES: readonly string[] = ['backpressure', 'protocol', 'replica', 'transport'];
/** What a received frame costs against `maxPendingBytes` beyond its bytes, so empty frames are charged too; a lone frame may still use the whole allowance. */
const FRAME_COST = 64;
/** How long received updates may wait for the provider's sync before they are posted anyway. */
const HELD_UPDATES_MS = 250;

/** Whether `data` is one sync Update message, which the worker's provider only holds until it has synced. */
function isSyncUpdate(data: Uint8Array): boolean {
  if (data[0] !== 0 || data[1] !== 2) return false;
  let length = 0;
  for (let index = 2, shift = 0; index < Math.min(data.byteLength, 7); index++, shift += 7) {
    length += (data[index]! & 0x7f) * 2 ** shift;
    if (data[index]! < 0x80) return index + 1 + length === data.byteLength;
  }
  return false;
}

/**
 * The main-thread side of a provider whose protocol runs in the document worker. It owns the host transport's
 * lifecycle as a provider does, forwards the transport's events with frame bytes untouched, one post per burst,
 * and sends what the worker answers. Worker output older than main's latest lifecycle call is dropped.
 */
export class PipedCollaboration<E extends Error, C = PipeCursor, P = PipePeer> {
  private readonly endpoint: CollaborationPipeEndpoint<C>;
  private readonly statusListeners = new Map<number, (change: { status: Status; synced: boolean }) => void>();
  private readonly errorListeners = new Map<number, (error: E) => void>();
  private readonly peerListeners = new Map<number, (peers: readonly P[]) => void>();
  private nextListenerId = 0;
  private connectionStatus: Status = 'disconnected';
  private isSynced = false;
  private peerList: readonly P[] = [];
  private seq = 0;
  /** The op of main's latest lifecycle call; older worker output describes a connection main has left. */
  private floor = 0;
  /** The op of the transport's latest open or close; sends, status and peers from before it describe a connection the host ended. */
  private sendFloor = 0;
  /** Received frame bytes queued or in flight to the worker. */
  private inboundBytes = 0;
  private flushingFrames = false;
  private queue: PipeOp<C>[] = [];
  private posting = false;
  private flushQueued = false;
  private heldTimer: ReturnType<typeof setTimeout> | null = null;
  /** The queue holds an op that should not wait for held updates. */
  private urgent = false;
  private unsubscribeTransport: (() => void) | null = null;
  private subscription = 0;
  private wantsConnection = false;
  private connectAttempt = 0;
  private leaveFrame: Uint8Array | null = null;
  private readonly pendingFrames: Uint8Array[] = [];
  private queuedBytes = 0;
  private isDestroyed = false;
  private closed = false;

  constructor(
    channel: CollaborationPipeChannel<C, P>,
    private readonly transport: PipeHostTransport,
    private readonly options: PipedCollaborationOptions<E>
  ) {
    this.endpoint = channel.open((out) => this.receive(out), (reason) => this.lose(reason));
    this.enqueue({ op: 'create', seq: ++this.seq, options: options.worker });
  }

  get status(): Status {
    return this.connectionStatus;
  }

  get synced(): boolean {
    return this.isSynced;
  }

  get pendingBytes(): number {
    return this.queuedBytes;
  }

  get peers(): readonly P[] {
    return this.peerList;
  }

  setCursor(cursor: C | null, broadcast: boolean): void {
    if (this.isDestroyed) return;
    this.enqueue({ op: 'cursor', seq: ++this.seq, cursor, broadcast });
  }

  onPeers(listener: (peers: readonly P[]) => void): () => void {
    if (this.isDestroyed) return () => {};
    const id = this.nextListenerId++;
    this.peerListeners.set(id, listener);
    listener(this.peers);
    return () => this.peerListeners.delete(id);
  }

  onStatus(listener: (change: { status: Status; synced: boolean }) => void): () => void {
    if (this.isDestroyed) return () => {};
    const id = this.nextListenerId++;
    this.statusListeners.set(id, listener);
    return () => this.statusListeners.delete(id);
  }

  onError(listener: (error: E) => void): () => void {
    if (this.isDestroyed) return () => {};
    const id = this.nextListenerId++;
    this.errorListeners.set(id, listener);
    return () => this.errorListeners.delete(id);
  }

  connect(): void {
    if (this.isDestroyed || this.connectionStatus === 'connected' || this.connectionStatus === 'connecting') return;
    if (this.closed) {
      this.report(this.options.makeError('replica', 'Collaboration replica is retired'));
      return;
    }
    this.wantsConnection = true;
    this.floor = ++this.seq;
    this.enqueue({ op: 'connect', seq: this.floor });
    const attempt = ++this.connectAttempt;
    this.setStatus('connecting', false);
    if (!this.wantsConnection || this.isDestroyed || attempt !== this.connectAttempt) return;

    if (!this.unsubscribeTransport) {
      const subscription = ++this.subscription;
      let unsubscribe: () => void;
      try {
        unsubscribe = this.transport.onEvent((event) => this.forward(subscription, event));
        if (typeof unsubscribe !== 'function') {
          throw new TypeError('Transport event subscription must return a function');
        }
      } catch (cause) {
        this.subscription += 1;
        this.fail(this.transportError('Failed to subscribe to transport events', cause));
        return;
      }
      if (!this.wantsConnection || this.isDestroyed || attempt !== this.connectAttempt || subscription !== this.subscription) {
        try {
          unsubscribe();
        } catch (cause) {
          if (!this.isDestroyed) this.report(this.transportError('Failed to unsubscribe from stale transport events', cause));
        }
        return;
      }
      this.unsubscribeTransport = unsubscribe;
    }

    try {
      const result = this.transport.connect();
      if (result) {
        void result.catch((cause) => {
          if (!this.isDestroyed && this.wantsConnection && this.connectAttempt === attempt) {
            this.fail(this.transportError('Transport connect failed', cause));
          }
        });
      }
    } catch (cause) {
      this.fail(this.transportError('Transport connect failed', cause));
    }
  }

  disconnect(): void {
    if (this.isDestroyed) return;
    if (this.closed) {
      this.setStatus('disconnected', false);
      return;
    }
    const active = this.wantsConnection || this.unsubscribeTransport !== null || this.connectionStatus !== 'disconnected';
    if (!active) return;
    this.sendLeave();
    this.floor = ++this.seq;
    this.enqueue({ op: 'disconnect', seq: this.floor });
    const errors = this.release(true);
    this.setStatus('disconnected', false);
    for (const error of errors) this.report(error);
  }

  destroy(): void {
    if (this.isDestroyed) return;
    const active = !this.closed && (this.wantsConnection || this.unsubscribeTransport !== null ||
      this.connectionStatus === 'connecting' || this.connectionStatus === 'connected');
    if (active) this.sendLeave();
    this.floor = ++this.seq;
    this.enqueue({ op: 'destroy', seq: this.floor });
    const errors = this.closed ? [] : this.release(active);
    this.isDestroyed = true;
    this.setStatus('destroyed', false);
    for (const error of errors) this.report(error);
    this.statusListeners.clear();
    this.errorListeners.clear();
    this.peerListeners.clear();
  }

  private transportError(context: string, cause: unknown): E {
    return this.options.normalize('transport', context, cause);
  }

  /** Leaves the connection main holds: drops the subscription, disconnects, clears queued bytes and peers. */
  private release(disconnect: boolean): E[] {
    this.wantsConnection = false;
    this.connectAttempt += 1;
    this.leaveFrame = null;
    this.clearQueue();
    this.clearPeers();
    const errors: E[] = [];
    const unsubscribe = this.unsubscribeTransport;
    this.unsubscribeTransport = null;
    this.subscription += 1;
    if (unsubscribe) {
      try {
        unsubscribe();
      } catch (cause) {
        errors.push(this.transportError('Failed to unsubscribe from transport events', cause));
      }
    }
    if (!disconnect) return errors;
    try {
      const result = this.transport.disconnect();
      if (result) {
        void result.catch((cause) => {
          if (!this.isDestroyed) this.report(this.transportError('Transport disconnect failed', cause));
        });
      }
    } catch (cause) {
      errors.push(this.transportError('Transport disconnect failed', cause));
    }
    return errors;
  }

  /** Fails main's connection with `error`, and the worker's with it. */
  private fail(error: E): void {
    if (this.isDestroyed || !this.wantsConnection) return;
    const code = (error as { code?: unknown }).code;
    this.floor = ++this.seq;
    this.enqueue({ op: 'fail', seq: this.floor, error: {
      code: (typeof code === 'string' && ERROR_CODES.includes(code) ? code : 'transport') as PipeError['code'],
      message: error.message,
    } });
    const errors = this.release(true);
    this.setStatus('disconnected', false);
    this.report(error);
    for (const cleanup of errors) this.report(cleanup);
  }

  private sendLeave(): void {
    const frame = this.leaveFrame;
    this.leaveFrame = null;
    if (!frame || !this.options.hasUser || !this.unsubscribeTransport || this.connectionStatus !== 'connected') return;
    try {
      this.transport.send(frame);
    } catch (cause) {
      this.report(this.transportError('Failed to send awareness leave', cause));
    }
  }

  private enqueue(op: PipeOp<C>, held = false): void {
    if (this.closed) return;
    this.queue.push(op);
    this.schedule(held);
  }

  private schedule(held: boolean): void {
    if (!held) this.urgent = true;
    if (this.posting || this.flushQueued) return;
    if (held) {
      this.heldTimer ??= setTimeout(() => { this.heldTimer = null; this.flush(); }, HELD_UPDATES_MS);
      return;
    }
    this.flushQueued = true;
    queueMicrotask(() => {
      this.flushQueued = false;
      this.flush();
    });
  }

  private flush(): void {
    if (this.posting || this.closed || this.queue.length === 0) return;
    if (this.heldTimer !== null) { clearTimeout(this.heldTimer); this.heldTimer = null; }
    const ops = this.queue;
    this.queue = [];
    this.urgent = false;
    this.posting = true;
    const destroying = ops.some((op) => op.op === 'destroy');
    const inbound = ops.reduce((sum, op) =>
      op.op === 'event' && op.event.type === 'message' ? sum + op.event.data.byteLength + FRAME_COST : sum, 0);
    let posted: Promise<unknown>;
    try {
      posted = this.endpoint.post(ops, pipeOpBuffers(ops));
    } catch (cause) {
      posted = Promise.reject(cause);
    }
    posted.then(() => {
      this.posting = false;
      this.inboundBytes -= inbound;
      if (destroying) this.close();
      else if (this.urgent) this.flush();
      else if (this.queue.length > 0) this.schedule(true);
    }, (cause: unknown) => {
      this.posting = false;
      this.inboundBytes -= inbound;
      this.lose(message(cause));
    });
  }

  private forward(subscription: number, event: TransportEvent): void {
    if (this.closed || this.isDestroyed || subscription !== this.subscription) return;
    switch (event.type) {
      case 'message':
        if (this.inboundBytes === 0
          ? event.data.byteLength > this.options.maxPendingBytes
          : event.data.byteLength + FRAME_COST > this.options.maxPendingBytes - this.inboundBytes) {
          this.fail(this.options.makeError('backpressure', 'Pending replica data exceeds the byte limit'));
          return;
        }
        this.inboundBytes += event.data.byteLength + FRAME_COST;
        this.enqueue({ op: 'event', seq: ++this.seq, event: { type: 'message', data: event.data.slice().buffer } },
          !this.isSynced && isSyncUpdate(event.data) && this.inboundBytes <= this.options.maxPendingBytes / 2);
        return;
      case 'open':
      case 'close':
        this.clearQueue();
        if (event.type === 'close') {
          this.connectAttempt += 1;
          this.leaveFrame = null;
        }
        this.sendFloor = ++this.seq;
        this.enqueue({ op: 'event', seq: this.sendFloor,
          event: event.type === 'open' ? { type: 'open' } : { type: 'close', ...(event.reason !== undefined && { reason: event.reason }) } });
        if (event.type === 'close') {
          this.clearPeers();
          this.setStatus('disconnected', false);
        }
        return;
      case 'drain':
        this.flushPending();
        return;
      case 'error':
        this.fail(this.transportError('Transport error', event.error));
        return;
      default:
        this.fail(this.options.makeError('transport', 'Unknown transport event'));
    }
  }

  private receive(out: readonly PipeOutput<P>[]): void {
    for (const output of out) {
      if (this.closed || this.isDestroyed) return;
      if (output.do === 'lost') {
        this.lose('Collaboration replica is retired');
        return;
      }
      if (output.through < this.floor) continue;
      switch (output.do) {
        case 'unsubscribe': {
          const errors = this.release(false);
          for (const error of errors) this.report(error);
          break;
        }
        case 'disconnect': {
          this.wantsConnection = false;
          const errors = this.release(true);
          for (const error of errors) this.report(error);
          break;
        }
        case 'send':
          if (output.through >= this.sendFloor) this.send(new Uint8Array(output.data));
          break;
        case 'leave':
          if (output.through >= this.sendFloor) this.leaveFrame = output.data ? new Uint8Array(output.data) : null;
          break;
        case 'status':
          if (output.through >= this.sendFloor) this.setStatus(output.status, output.synced);
          break;
        case 'peers':
          if (output.through < this.sendFloor || (output.peers.length === 0 && this.peerList.length === 0)) break;
          this.peerList = output.peers;
          this.emitPeers();
          break;
        case 'error':
          this.report(this.options.makeError(output.error.code, output.error.message));
          break;
      }
    }
  }

  private send(frame: Uint8Array): void {
    if (!this.unsubscribeTransport) return;
    if (this.pendingFrames.length > 0) {
      this.queueFrame(frame);
      return;
    }
    const connection = this.sendFloor;
    let accepted: unknown;
    try {
      accepted = this.transport.send(frame.slice());
    } catch (cause) {
      this.fail(this.transportError('Transport send failed', cause));
      return;
    }
    if (typeof accepted !== 'boolean') {
      this.fail(this.options.makeError('transport', 'Transport send must return a boolean'));
      return;
    }
    if (!accepted && this.unsubscribeTransport !== null && connection === this.sendFloor) this.queueFrame(frame);
  }

  private queueFrame(frame: Uint8Array): void {
    if (frame.byteLength > this.options.maxPendingBytes - this.queuedBytes) {
      this.fail(this.options.makeError('backpressure', `Pending collaboration data exceeds ${this.options.maxPendingBytes} bytes`));
      return;
    }
    this.pendingFrames.push(frame);
    this.queuedBytes += frame.byteLength;
  }

  private flushPending(): void {
    if (this.flushingFrames) return;
    this.flushingFrames = true;
    try {
      while (this.unsubscribeTransport && this.pendingFrames.length > 0) {
        const frame = this.pendingFrames[0]!;
        let accepted: unknown;
        try {
          accepted = this.transport.send(frame.slice());
        } catch (cause) {
          this.fail(this.transportError('Transport send failed', cause));
          return;
        }
        if (typeof accepted !== 'boolean') {
          this.fail(this.options.makeError('transport', 'Transport send must return a boolean'));
          return;
        }
        if (!accepted || this.pendingFrames[0] !== frame) return;
        this.pendingFrames.shift();
        this.queuedBytes -= frame.byteLength;
      }
    } finally {
      this.flushingFrames = false;
    }
  }

  private clearQueue(): void {
    this.pendingFrames.length = 0;
    this.queuedBytes = 0;
  }

  /**
   * The worker no longer runs this provider. Hands the host transport back at once; reports the loss a task later,
   * so a host that destroys the provider on `onReplica(null)` hears nothing.
   */
  private lose(reason: string): void {
    if (this.closed) return;
    this.close();
    if (this.isDestroyed) return;
    this.sendLeave();
    const errors = this.release(this.wantsConnection || this.connectionStatus !== 'disconnected');
    setTimeout(() => {
      if (this.isDestroyed) return;
      this.setStatus('disconnected', false);
      this.report(this.options.makeError('replica', reason));
      for (const error of errors) this.report(error);
    }, 0);
  }

  private close(): void {
    if (this.closed) return;
    this.closed = true;
    this.queue = [];
    if (this.heldTimer !== null) { clearTimeout(this.heldTimer); this.heldTimer = null; }
    this.endpoint.close();
  }

  private clearPeers(): void {
    if (this.peerList.length === 0) return;
    this.peerList = [];
    this.emitPeers();
  }

  private setStatus(status: Status, synced: boolean): void {
    if (this.connectionStatus === status && this.isSynced === synced) return;
    this.connectionStatus = status;
    if (synced && !this.isSynced && this.queue.length > 0) this.schedule(false);
    this.isSynced = synced;
    for (const [id, listener] of [...this.statusListeners]) {
      if (this.statusListeners.get(id) !== listener) continue;
      try { listener({ status, synced }); } catch {}
    }
  }

  private emitPeers(): void {
    const peers = this.peers;
    for (const [id, listener] of [...this.peerListeners]) {
      if (this.peerListeners.get(id) !== listener) continue;
      try { listener(peers); } catch {}
    }
  }

  private report(error: E): void {
    for (const [id, listener] of [...this.errorListeners]) {
      if (this.errorListeners.get(id) !== listener) continue;
      try { listener(error); } catch {}
    }
  }
}

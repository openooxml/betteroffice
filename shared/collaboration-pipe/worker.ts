import {
  clockOrigin,
  ownedBuffer,
  type PipeCommandPayload,
  type PipeCursor,
  type PipeError,
  type PipeOutput,
  type PipeOutputBatch,
  type PipePeer,
  type PipeProviderOptions,
} from './protocol';

type UpdateOrigin = 'local' | 'remote';
type TransportEvent =
  | { type: 'open' }
  | { type: 'message'; data: Uint8Array }
  | { type: 'close'; reason?: string }
  | { type: 'error'; error: unknown };

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
  onStatus(listener: (change: { status: Extract<PipeOutput, { do: 'status' }>['status']; synced: boolean }) => void): () => void;
  onError(listener: (error: { code: string; message: string }) => void): () => void;
  onPeers(listener: (peers: readonly P[]) => void): () => void;
}

/** A merge the provider handed over: settled once the burst's updates merge, failed when they do not. */
export interface PipeArrival {
  settle(applied: unknown): void;
  fail(cause: unknown): void;
}

export type PipeMergeResult = { applied: unknown } | { failed: unknown };

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
  leave(provider: Prov): { clock: number; frame: Uint8Array } | null;
  /** The error object a transport `error` event carries for `error`. */
  transportError(error: PipeError): unknown;
  /** Output became pending outside `run`, from the providers' timers. */
  onOutput?(): void;
}

export interface CollaborationPipeHost<C = PipeCursor, P = PipePeer> {
  /** Runs one burst of ops: transport events in order, then one merge of every update they carried. */
  run(payload: PipeCommandPayload<C>): void;
  /** Pending output, in order, and the buffers it holds. */
  drain(): { batches: PipeOutputBatch<P>[]; transfer: ArrayBuffer[] } | undefined;
  /** Forgets every provider without sending anything. */
  dispose(): void;
}

interface Endpoint<Prov> {
  id: number;
  provider: Prov;
  transport: WorkerPipeTransport;
  through: number;
  /** While true the worker replays a call main already made: it sends nothing and asks for no transport call. */
  echo: boolean;
  leaveClock: number | null;
  stop: Array<() => void>;
}

/** The provider's transport inside the worker: main subscribes and connects; this only routes and reports. */
class WorkerPipeTransport implements PipeTransport {
  private listener: ((event: TransportEvent) => void) | null = null;

  constructor(private readonly emit: (output: 'unsubscribe' | 'disconnect' | Uint8Array) => void) {}

  connect(): void {}

  disconnect(): void {
    this.emit('disconnect');
  }

  send(data: Uint8Array): boolean {
    this.emit(data);
    return true;
  }

  onEvent(listener: (event: TransportEvent) => void): () => void {
    this.listener = listener;
    return () => {
      if (this.listener !== listener) return;
      this.listener = null;
      this.emit('unsubscribe');
    };
  }

  deliver(event: TransportEvent): void {
    this.listener?.(event);
  }
}

export function createCollaborationPipeHost<Prov extends PipeProvider<C, P>, C = PipeCursor, P = PipePeer>(
  options: CollaborationPipeHostOptions<Prov, C, P>
): CollaborationPipeHost<C, P> {
  const endpoints = new Map<number, Endpoint<Prov>>();
  let outbox: PipeOutputBatch<P>[] = [];
  let running = false;
  let disposed = false;
  let burst: { updates: Uint8Array[]; arrivals: PipeArrival[]; single: boolean } | null = null;

  const emit = (provider: number, output: PipeOutput<P>) => {
    if (disposed) return;
    const notify = !running && outbox.length === 0;
    const last = outbox[outbox.length - 1];
    if (last?.provider === provider) last.out.push(output);
    else outbox.push({ provider, out: [output] });
    if (notify) options.onOutput?.();
  };

  const flush = () => {
    const pending = burst;
    if (!pending || pending.updates.length === 0) return;
    burst = { updates: [], arrivals: [], single: true };
    const result = options.merge(pending.updates, pending.single && pending.updates.length === 1);
    if ('failed' in result) {
      for (const arrival of pending.arrivals) arrival.fail(result.failed);
      return;
    }
    pending.arrivals.forEach((arrival, index) =>
      arrival.settle(index === pending.arrivals.length - 1 ? result.applied : null));
  };

  const hold = (provider: Prov, updates: readonly Uint8Array[], single: boolean) => {
    if (!burst) throw new Error('Collaboration updates merge only within a pipe burst');
    burst.updates.push(...updates);
    burst.arrivals.push(options.arrive(provider));
    burst.single &&= single;
    return null;
  };

  const replicaFor = (provider: () => Prov): PipeReplica => ({
    clientId: options.clientId,
    encodeStateVector: () => { flush(); return options.read.encodeStateVector(); },
    encodeStateAsUpdate: (vector) => { flush(); return options.read.encodeStateAsUpdate(vector); },
    ...(options.read.encodeStateAsUpdateChunks && {
      encodeStateAsUpdateChunks: (vector: Uint8Array, maxChunkBytes: number) => {
        flush();
        return options.read.encodeStateAsUpdateChunks!(vector, maxChunkBytes);
      },
    }),
    applyUpdate: (update) => hold(provider(), [update], true),
    applyUpdateChunks: (updates) => hold(provider(), updates, false),
    onUpdate: options.onUpdate,
  });

  const create = (id: number, providerOptions: PipeProviderOptions): Endpoint<Prov> => {
    let endpoint!: Endpoint<Prov>;
    const transport = new WorkerPipeTransport((output) => {
      if (endpoint.echo) return;
      emit(id, typeof output === 'string'
        ? { do: output, through: endpoint.through }
        : { do: 'send', data: ownedBuffer(output), through: endpoint.through });
    });
    const provider = options.createProvider(replicaFor(() => endpoint.provider), transport, providerOptions);
    endpoint = { id, provider, transport, through: 0, echo: false, leaveClock: null, stop: [] };
    let subscribed = false;
    endpoint.stop.push(
      provider.onStatus(({ status, synced }) => emit(id, { do: 'status', status, synced, through: endpoint.through })),
      provider.onError((error) => {
        if (!endpoint.echo) {
          emit(id, { do: 'error', error: { code: error.code as PipeError['code'], message: error.message }, through: endpoint.through });
        }
      }),
      provider.onPeers((peers) => {
        if (subscribed) emit(id, { do: 'peers', peers: [...peers], timeOrigin: clockOrigin(),
          through: endpoint.through });
      }),
    );
    subscribed = true;
    endpoints.set(id, endpoint);
    return endpoint;
  };

  const echo = (endpoint: Endpoint<Prov>, call: () => void) => {
    endpoint.echo = true;
    try { call(); } finally { endpoint.echo = false; }
  };

  const run = (payload: PipeCommandPayload<C>) => {
    let endpoint = endpoints.get(payload.provider);
    for (const op of payload.ops) {
      if (op.op === 'create') {
        endpoint ??= create(payload.provider, op.options);
        endpoint.through = op.seq;
        continue;
      }
      if (!endpoint) {
        emit(payload.provider, { do: 'lost' });
        return;
      }
      const current = endpoint;
      current.through = op.seq;
      switch (op.op) {
        case 'connect':
          echo(current, () => current.provider.connect());
          break;
        case 'disconnect':
          echo(current, () => current.provider.disconnect());
          break;
        case 'fail':
          echo(current, () => current.transport.deliver({ type: 'error', error: options.transportError(op.error) }));
          break;
        case 'destroy':
          echo(current, () => current.provider.destroy());
          for (const stop of current.stop) stop();
          endpoints.delete(current.id);
          endpoint = undefined;
          break;
        case 'cursor':
          current.provider.setCursor(op.cursor, op.broadcast);
          break;
        case 'event': {
          const { event } = op;
          current.transport.deliver(event.type === 'message' ? { type: 'message', data: new Uint8Array(event.data) } : event);
          break;
        }
      }
    }
  };

  const noteLeaves = () => {
    for (const endpoint of endpoints.values()) {
      const leave = options.leave(endpoint.provider);
      const clock = leave?.clock ?? null;
      if (clock === endpoint.leaveClock) continue;
      endpoint.leaveClock = clock;
      emit(endpoint.id, { do: 'leave', data: leave ? ownedBuffer(leave.frame) : null, through: endpoint.through });
    }
  };

  return {
    run(payload) {
      if (disposed) return;
      running = true;
      burst = { updates: [], arrivals: [], single: true };
      try {
        run(payload);
        flush();
      } finally {
        burst = null;
        running = false;
      }
    },
    drain() {
      running = true;
      try { noteLeaves(); } finally { running = false; }
      if (outbox.length === 0) return undefined;
      const batches = outbox;
      outbox = [];
      const transfer: ArrayBuffer[] = [];
      for (const batch of batches) {
        for (const output of batch.out) {
          if (output.do === 'send') transfer.push(output.data);
          else if (output.do === 'leave' && output.data) transfer.push(output.data);
        }
      }
      return { batches, transfer };
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      for (const endpoint of endpoints.values()) {
        for (const stop of endpoint.stop) stop();
        endpoint.echo = true;
        try { endpoint.provider.destroy(); } catch {}
      }
      endpoints.clear();
      outbox = [];
    },
  };
}

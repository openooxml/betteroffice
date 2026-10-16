import { afterEach, describe, expect, it } from 'bun:test';
import { createAsyncCollaborationReplica } from './asyncReplica';
import {
  CHUNKED_SYNC_CLIENT,
  decodeMessages,
  encodeAwarenessMessage,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
  encodeVarUint,
  MAX_FRAME_BYTES,
  requestsChunkedSync,
  withChunkedSyncRequest,
} from './protocol';
import { CollaborationProvider } from './provider';
import type {
  AsyncCollaborationReplica,
  CollaborationError,
  CollaborationProviderOptions,
  CollaborationReplica,
  CollaborationTransport,
  CollaborationTransportEvent,
  CollaborationUpdateOrigin,
} from './types';

const MiB = 1024 * 1024;
const inference = { clientId: 8, story: 'body', paraId: 'p1', endOffset: 6 };

class ChunkReplica implements CollaborationReplica {
  readonly clientId = 1;
  whole = new Uint8Array(64);
  chunks: unknown = [Uint8Array.of(1), Uint8Array.of(2), Uint8Array.of(0, 3)];
  readonly limits: number[] = [];
  readonly applied: Uint8Array[] = [];
  readonly syncs: Uint8Array[][] = [];
  applyError: unknown;

  encodeStateVector(): Uint8Array {
    return Uint8Array.of(0);
  }

  encodeStateAsUpdate(): Uint8Array {
    return this.whole;
  }

  encodeStateAsUpdateChunks(_vector: Uint8Array, maxChunkBytes: number): Uint8Array[] {
    this.limits.push(maxChunkBytes);
    return this.chunks as Uint8Array[];
  }

  applyUpdate(update: Uint8Array): unknown {
    this.applied.push(update);
    return inference;
  }

  applyUpdateChunks(updates: readonly Uint8Array[]): void {
    if (this.applyError) throw this.applyError;
    this.syncs.push([...updates]);
  }

  onUpdate(_listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void {
    return () => {};
  }
}

class Transport implements CollaborationTransport {
  readonly sent: Uint8Array[] = [];
  accept = true;
  onSend: (() => void) | undefined;
  disconnectCount = 0;
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();

  connect(): void {}

  disconnect(): void {
    this.disconnectCount += 1;
  }

  send(frame: Uint8Array): boolean {
    if (!this.accept) return false;
    this.sent.push(frame.slice());
    this.onSend?.();
    return true;
  }

  onEvent(listener: (event: CollaborationTransportEvent) => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  emit(event: CollaborationTransportEvent): void {
    for (const listener of [...this.listeners]) listener(event);
  }

  receive(data: Uint8Array): void {
    this.emit({ type: 'message', data });
  }
}

const providers = new Set<CollaborationProvider>();

afterEach(() => {
  for (const provider of providers) provider.destroy();
  providers.clear();
});

function open(replica: CollaborationReplica, options: CollaborationProviderOptions = {}) {
  const transport = new Transport();
  const provider = new CollaborationProvider(replica, transport, options);
  providers.add(provider);
  const errors: CollaborationError[] = [];
  provider.onError((error) => errors.push(error));
  provider.connect();
  transport.emit({ type: 'open' });
  transport.sent.length = 0;
  return { provider, transport, errors };
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((bytes, part) => bytes + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.byteLength;
  }
  return output;
}

/** The replica without `applyUpdateChunks` or `encodeStateAsUpdateChunks`, as a peer from before them. */
function plainReplica(replica: ChunkReplica): CollaborationReplica {
  return {
    clientId: replica.clientId,
    encodeStateVector: () => replica.encodeStateVector(),
    encodeStateAsUpdate: () => replica.encodeStateAsUpdate(),
    applyUpdate: (update) => replica.applyUpdate(update),
    onUpdate: (listener) => replica.onUpdate(listener),
  };
}

function peerAwareness(clock = 1): Uint8Array {
  return encodeAwarenessMessage([{
    clientId: 8,
    clock,
    state: { user: { name: 'Bright Fox', color: '#137333' }, cursor: null },
  }]);
}

function sent(transport: Transport) {
  return transport.sent.flatMap((frame) => decodeMessages(frame));
}

const asking = () => encodeSyncStep1(withChunkedSyncRequest(Uint8Array.of(0)));
const chunked = { syncChunkBytes: MiB };

describe('chunked sync requests in the state vector', () => {
  it('adds one empty marker client to a well-formed vector and nothing otherwise', () => {
    const vector = Uint8Array.of(1, 5, 9);
    const marked = withChunkedSyncRequest(vector);
    expect(requestsChunkedSync(vector)).toBe(false);
    expect(requestsChunkedSync(marked)).toBe(true);
    expect([...marked]).toEqual([2, 5, 9, ...encodeVarUint(CHUNKED_SYNC_CLIENT), 0]);
    expect(withChunkedSyncRequest(marked)).toBe(marked);
    for (const malformed of [Uint8Array.of(4), Uint8Array.of(1, 5), new Uint8Array()]) {
      expect(withChunkedSyncRequest(malformed)).toBe(malformed);
      expect(requestsChunkedSync(malformed)).toBe(false);
    }
    expect(requestsChunkedSync(Uint8Array.of(1, ...encodeVarUint(CHUNKED_SYNC_CLIENT), 3))).toBe(false);
  });

  it('asks for chunks only when the replica applies them as one change', () => {
    const handshakes = [new ChunkReplica(), plainReplica(new ChunkReplica())].map((replica) => {
      const transport = new Transport();
      const provider = new CollaborationProvider(replica, transport);
      providers.add(provider);
      provider.connect();
      transport.emit({ type: 'open' });
      return sent(transport)[0];
    });
    expect(handshakes).toEqual([
      { type: 'sync-step-1', stateVector: withChunkedSyncRequest(Uint8Array.of(1, 1, 1)) },
      { type: 'sync-step-1', stateVector: Uint8Array.of(1, 1, 1) },
    ]);
  });
});

describe('CollaborationProvider chunked sync answers', () => {
  it('answers a SyncStep1 that asks for chunks with Update chunks closed by a SyncStep2', () => {
    const replica = new ChunkReplica();
    const { transport, errors } = open(replica, chunked);
    transport.receive(asking());

    expect(sent(transport)).toEqual([
      { type: 'update', update: Uint8Array.of(1) },
      { type: 'update', update: Uint8Array.of(2) },
      { type: 'sync-step-2', update: Uint8Array.of(0, 3) },
    ]);
    expect(replica.limits).toEqual([MiB]);
    expect(errors).toEqual([]);
  });

  it('answers a SyncStep1 that does not ask with the whole diff as one SyncStep2', () => {
    const replica = new ChunkReplica();
    const { transport } = open(replica);
    transport.receive(encodeSyncStep1(Uint8Array.of(0)));
    expect(sent(transport)).toEqual([{ type: 'sync-step-2', update: replica.whole }]);
    expect(replica.limits).toEqual([]);
  });

  it('keeps chunks within syncChunkBytes and one frame', () => {
    const limits = [
      { syncChunkBytes: 4096 },
      { maxFrameBytes: 64, syncChunkBytes: 64 },
      { maxFrameBytes: 4096, syncChunkBytes: 4096 },
    ].map((options) => {
      const replica = new ChunkReplica();
      open(replica, options).transport.receive(asking());
      return replica.limits;
    });
    expect(limits).toEqual([[4096], [64 - 3], [4096 - 4]]);
  });

  it('fails before sending when a chunk does not fit a frame', () => {
    const replica = new ChunkReplica();
    replica.chunks = [Uint8Array.of(1), new Uint8Array(64)];
    const { provider, transport, errors } = open(replica, { maxFrameBytes: 32, syncChunkBytes: 32 });
    transport.receive(asking());

    expect(transport.sent).toEqual([]);
    expect(errors.map((error) => [error.code, error.message])).toEqual([
      ['protocol', 'Failed to encode SyncStep2: Frame exceeds 32 bytes'],
    ]);
    expect(provider.status).toBe('disconnected');
  });

  for (const chunks of [[], [1, 2], null, [Uint8Array.of(1), new Uint8Array()]]) {
    it(`fails with replica when the chunks are ${JSON.stringify(chunks)}`, () => {
      const replica = new ChunkReplica();
      replica.chunks = chunks;
      const { transport, errors } = open(replica, chunked);
      transport.receive(asking());

      expect(transport.sent).toEqual([]);
      expect(errors.map((error) => error.code)).toEqual(['replica']);
    });
  }

  it('stops an answer when the transport reopens while sending it', () => {
    const replica = new ChunkReplica();
    const { transport, errors } = open(replica, chunked);
    transport.onSend = () => {
      transport.onSend = undefined;
      transport.emit({ type: 'close' });
      transport.emit({ type: 'open' });
    };
    transport.receive(asking());

    expect(sent(transport)).toEqual([
      { type: 'update', update: Uint8Array.of(1) },
      { type: 'sync-step-1', stateVector: withChunkedSyncRequest(Uint8Array.of(1, 1, 1)) },
    ]);
    expect(errors).toEqual([]);
  });

  it('queues refused chunks in order against maxPendingBytes', () => {
    const replica = new ChunkReplica();
    const { provider, transport, errors } = open(replica, chunked);
    transport.accept = false;
    transport.receive(asking());
    expect(provider.pendingBytes).toBeGreaterThan(0);
    transport.accept = true;
    transport.emit({ type: 'drain' });
    expect(sent(transport).map((message) => message.type)).toEqual(['update', 'update', 'sync-step-2']);
    expect(errors).toEqual([]);

    const tight = open(new ChunkReplica(), { ...chunked, maxPendingBytes: 8 });
    tight.transport.accept = false;
    tight.transport.receive(asking());
    expect(tight.errors.map((error) => error.code)).toEqual(['backpressure']);
    expect(tight.provider.status).toBe('disconnected');
  });
});

/** Delivers every frame a peer sends to all the others, as a relay room does, until none are left. */
function room(...transports: Transport[]): () => void {
  const delivered = transports.map((transport) => transport.sent.length);
  return () => {
    for (let moved = true; moved;) {
      moved = false;
      transports.forEach((from, index) => {
        while (delivered[index] < from.sent.length) {
          const frame = from.sent[delivered[index]++];
          moved = true;
          for (const to of transports) if (to !== from) to.receive(frame);
        }
      });
    }
  };
}

function shape(transport: Transport) {
  return sent(transport).map((message) =>
    message.type === 'update' || message.type === 'sync-step-2'
      ? [message.type, message.update.byteLength]
      : [message.type]
  );
}

const sizes = (updates: readonly Uint8Array[]) => updates.map((update) => update.byteLength);

describe('CollaborationProvider chunked sync answers by transport', () => {
  const whole = new Uint8Array(MiB + 4096).fill(7);

  function syncedPeer(replica: ChunkReplica, options: CollaborationProviderOptions = {}, plain = false) {
    const peer = open(plain ? plainReplica(replica) : replica, options);
    peer.transport.receive(encodeSyncStep2(Uint8Array.of(0)));
    replica.applied.length = 0;
    peer.transport.sent.length = 0;
    return peer;
  }

  function answerer(): ChunkReplica {
    const replica = new ChunkReplica();
    replica.whole = whole;
    replica.chunks = [whole.subarray(0, MiB), whole.subarray(MiB)];
    return replica;
  }

  function joinRoom(replica: ChunkReplica, ...peers: Transport[]) {
    const transport = new Transport();
    const provider = new CollaborationProvider(replica, transport);
    providers.add(provider);
    const flush = room(...peers, transport);
    provider.connect();
    transport.emit({ type: 'open' });
    flush();
    return provider;
  }

  it('answers with one SyncStep2 by default, so other peers in a room never apply part of it', () => {
    const stepwise = new ChunkReplica();
    const source = answerer();
    const joiner = new ChunkReplica();
    const a = syncedPeer(stepwise, {}, true);
    const b = syncedPeer(source);
    const provider = joinRoom(joiner, a.transport, b.transport);

    expect(shape(b.transport)).toEqual([['sync-step-2', whole.byteLength], ['sync-step-1'], ['sync-step-2', whole.byteLength]]);
    expect(source.limits).toEqual([]);
    expect(sizes(stepwise.applied))
      .toEqual([whole.byteLength, whole.byteLength, joiner.whole.byteLength, joiner.whole.byteLength]);
    expect(sizes(joiner.applied))
      .toEqual([stepwise.whole.byteLength, whole.byteLength, whole.byteLength, stepwise.whole.byteLength]);
    expect(joiner.syncs).toEqual([]);
    expect(provider.synced).toBe(true);
    expect([...a.errors, ...b.errors]).toEqual([]);
  });

  it('chunks for a holding joiner when syncChunkBytes is set on a point-to-point pair', () => {
    const source = answerer();
    const joiner = new ChunkReplica();
    const b = syncedPeer(source, { syncChunkBytes: MiB });
    const provider = joinRoom(joiner, b.transport);

    expect(shape(b.transport)).toEqual([['update', MiB], ['sync-step-2', 4096], ['sync-step-1']]);
    expect(source.limits).toEqual([MiB]);
    expect(joiner.syncs.map(sizes)).toEqual([[MiB, 4096]]);
    expect(joiner.applied).toEqual([]);
    expect(provider.synced).toBe(true);
    expect(b.errors).toEqual([]);
  });
});

describe('CollaborationProvider held initial sync', () => {
  it('holds updates in arrival order until the SyncStep2 and applies them with it as one change', () => {
    const replica = new ChunkReplica();
    const { provider, transport, errors } = open(replica, { user: { name: 'Calm Otter' } });
    transport.receive(peerAwareness());
    transport.receive(encodeUpdate(Uint8Array.of(1)));
    transport.receive(concat(encodeUpdate(Uint8Array.of(9)), peerAwareness(2)));
    transport.receive(encodeUpdate(Uint8Array.of(2)));
    expect(replica.applied).toEqual([]);
    expect(replica.syncs).toEqual([]);
    expect(provider.synced).toBe(false);

    transport.receive(encodeSyncStep2(Uint8Array.of(0, 3)));
    expect(replica.syncs).toEqual([
      [Uint8Array.of(1), Uint8Array.of(9), Uint8Array.of(2), Uint8Array.of(0, 3)],
    ]);
    expect(replica.applied).toEqual([]);
    expect(provider.synced).toBe(true);
    expect(provider.peers[0].inferredCursor).toBeNull();

    transport.receive(encodeUpdate(Uint8Array.of(4)));
    expect(replica.applied).toEqual([Uint8Array.of(4)]);
    expect(provider.peers[0].inferredCursor).toEqual(inference);
    transport.receive(encodeSyncStep2(Uint8Array.of(5)));
    expect(replica.applied).toEqual([Uint8Array.of(4), Uint8Array.of(5)]);
    expect(replica.syncs).toHaveLength(1);
    expect(errors).toEqual([]);
  });

  it('applies a lone SyncStep2 as an ordinary update', () => {
    const replica = new ChunkReplica();
    const { provider, transport } = open(replica);
    transport.receive(encodeSyncStep2(Uint8Array.of(3)));
    expect(replica.applied).toEqual([Uint8Array.of(3)]);
    expect(replica.syncs).toEqual([]);
    expect(provider.synced).toBe(true);
  });

  it('fails with backpressure once held updates and their overhead pass maxSyncBytes', () => {
    const replica = new ChunkReplica();
    const { provider, transport, errors } = open(replica, { maxSyncBytes: 2 * 66 + 1 });
    transport.receive(encodeUpdate(Uint8Array.of(1, 2)));
    transport.receive(encodeUpdate(Uint8Array.of(3, 4)));
    expect(errors).toEqual([]);
    transport.receive(encodeUpdate(Uint8Array.of(5, 6)));
    transport.receive(encodeSyncStep2(Uint8Array.of(0)));

    expect(errors.map((error) => [error.code, error.message])).toEqual([
      ['backpressure', `Initial sync exceeds ${2 * 66 + 1} bytes`],
    ]);
    expect(provider.status).toBe('disconnected');
    expect(replica.syncs).toEqual([]);
    expect(replica.applied).toEqual([]);
  });

  it('charges empty held updates their overhead', () => {
    const replica = new ChunkReplica();
    const { transport, errors } = open(replica, { maxSyncBytes: 64 * 3 });
    const empty = encodeUpdate(new Uint8Array());
    transport.receive(concat(empty, empty, empty));
    expect(errors).toEqual([]);
    transport.receive(empty);
    expect(errors.map((error) => error.code)).toEqual(['backpressure']);
  });

  it('fails with replica when the held sync does not apply', () => {
    const replica = new ChunkReplica();
    replica.applyError = new Error('bad update');
    const { provider, transport, errors } = open(replica);
    transport.receive(encodeUpdate(Uint8Array.of(1)));
    transport.receive(encodeSyncStep2(Uint8Array.of(0)));

    expect(errors.map((error) => error.code)).toEqual(['replica']);
    expect(provider.synced).toBe(false);
    expect(provider.status).toBe('disconnected');
  });

  it('drops what a closed connection held and holds again after reopening', () => {
    const replica = new ChunkReplica();
    const { transport, errors } = open(replica, { maxSyncBytes: 66 });
    transport.receive(encodeUpdate(Uint8Array.of(1, 2)));
    transport.emit({ type: 'close' });
    transport.emit({ type: 'open' });
    transport.receive(encodeUpdate(Uint8Array.of(3, 4)));
    transport.receive(encodeSyncStep2(Uint8Array.of(0)));
    expect(replica.syncs).toEqual([[Uint8Array.of(3, 4), Uint8Array.of(0)]]);
    expect(errors).toEqual([]);
  });

  it('sends one SyncStep1 per hold, after the SyncStep2, for the requests answered while holding', () => {
    const replica = new ChunkReplica();
    const { transport, errors } = open(replica, chunked);
    const ahead = encodeSyncStep1(withChunkedSyncRequest(Uint8Array.of(1, 5, 1)));
    transport.receive(ahead);
    transport.receive(ahead);
    transport.receive(encodeUpdate(Uint8Array.of(7)));
    expect(sent(transport).filter((message) => message.type === 'sync-step-1')).toEqual([]);

    transport.receive(encodeSyncStep2(Uint8Array.of(0)));
    transport.receive(encodeSyncStep2(Uint8Array.of(0)));
    expect(sent(transport).filter((message) => message.type === 'sync-step-1'))
      .toEqual([{ type: 'sync-step-1', stateVector: Uint8Array.of(0) }]);
    expect(replica.syncs).toHaveLength(1);
    expect(errors).toEqual([]);
  });

  it('applies updates on arrival for a replica without applyUpdateChunks', () => {
    const replica = new ChunkReplica();
    const { provider, transport } = open(plainReplica(replica));
    transport.receive(encodeUpdate(Uint8Array.of(1)));
    expect(replica.applied).toEqual([Uint8Array.of(1)]);
    transport.receive(encodeSyncStep2(Uint8Array.of(2)));
    expect(replica.applied).toEqual([Uint8Array.of(1), Uint8Array.of(2)]);
    expect(provider.synced).toBe(true);
  });
});

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (cause: unknown) => void;
  const promise = new Promise<T>((done, fail) => {
    resolve = done;
    reject = fail;
  });
  return { promise, resolve, reject };
}

async function settle(): Promise<void> {
  for (let turn = 0; turn < 20; turn++) await Promise.resolve();
}

class AsyncChunkReplica implements AsyncCollaborationReplica {
  readonly clientId = 1;
  readonly answers: { limit: number; result: ReturnType<typeof deferred<Uint8Array[]>> }[] = [];
  readonly merges: Uint8Array[] = [];
  readonly syncs: { updates: Uint8Array[]; result: ReturnType<typeof deferred<unknown>> }[] = [];

  encodeStateVector(): Promise<Uint8Array> {
    return Promise.resolve(Uint8Array.of(0));
  }

  encodeStateAsUpdate(): Promise<Uint8Array> {
    return Promise.resolve(new Uint8Array(64));
  }

  encodeStateAsUpdateChunks(_vector: Uint8Array, limit: number): Promise<Uint8Array[]> {
    const result = deferred<Uint8Array[]>();
    this.answers.push({ limit, result });
    return result.promise;
  }

  applyUpdate(update: Uint8Array): Promise<unknown> {
    this.merges.push(update);
    return Promise.resolve(inference);
  }

  applyUpdateChunks(updates: readonly Uint8Array[]): Promise<unknown> {
    const result = deferred<unknown>();
    this.syncs.push({ updates: [...updates], result });
    return result.promise;
  }

  onUpdate(): () => void {
    return () => {};
  }
}

async function openAsync(options: CollaborationProviderOptions = {}) {
  const replica = new AsyncChunkReplica();
  const value = open(createAsyncCollaborationReplica(replica), options);
  await settle();
  value.transport.sent.length = 0;
  return { replica, ...value };
}

describe('CollaborationProvider chunked sync with async replicas', () => {
  it('answers SyncStep1 with chunks from the async replica', async () => {
    const { replica, transport, errors } = await openAsync(chunked);
    transport.receive(asking());
    await settle();
    expect(replica.answers.map((answer) => answer.limit)).toEqual([MiB]);
    replica.answers[0].result.resolve([Uint8Array.of(1), Uint8Array.of(0, 2)]);
    await settle();

    expect(sent(transport)).toEqual([
      { type: 'update', update: Uint8Array.of(1) },
      { type: 'sync-step-2', update: Uint8Array.of(0, 2) },
    ]);
    expect(errors).toEqual([]);
  });

  it('holds updates past the inbound limit and merges them with the SyncStep2 in order', async () => {
    const { replica, provider, transport, errors } = await openAsync({ ...chunked, maxPendingBytes: 24 });
    transport.receive(asking());
    for (let value = 1; value <= 4; value++) {
      transport.receive(encodeUpdate(Uint8Array.of(value, value, value)));
      transport.receive(peerAwareness(value));
    }
    transport.receive(encodeSyncStep2(Uint8Array.of(0, 9)));
    transport.receive(encodeUpdate(Uint8Array.of(5)));
    await settle();
    expect(replica.syncs).toEqual([]);
    expect(replica.merges).toEqual([]);

    replica.answers[0].result.resolve([Uint8Array.of(0)]);
    await settle();
    expect(replica.syncs.map((sync) => sync.updates)).toEqual([[
      Uint8Array.of(1, 1, 1), Uint8Array.of(2, 2, 2), Uint8Array.of(3, 3, 3), Uint8Array.of(4, 4, 4),
      Uint8Array.of(0, 9),
    ]]);
    expect(provider.synced).toBe(false);
    expect(replica.merges).toEqual([]);

    replica.syncs[0].result.resolve(inference);
    await settle();
    expect(provider.synced).toBe(true);
    expect(replica.merges).toEqual([Uint8Array.of(5)]);
    expect(errors).toEqual([]);
  });

  it('sends one SyncStep1 per hold, after the SyncStep2 merges, for the requests that arrived while holding', async () => {
    const { replica, transport, errors } = await openAsync(chunked);
    const requests = () => sent(transport).filter((message) => message.type === 'sync-step-1');
    const ahead = encodeSyncStep1(withChunkedSyncRequest(Uint8Array.of(1, 5, 1)));
    transport.receive(ahead);
    transport.receive(ahead);
    transport.receive(encodeUpdate(Uint8Array.of(7)));
    transport.receive(encodeSyncStep2(Uint8Array.of(0)));
    for (let answer = 0; answer < 2; answer++) {
      await settle();
      replica.answers[answer].result.resolve([Uint8Array.of(0)]);
    }
    await settle();
    expect(replica.syncs).toHaveLength(1);
    expect(requests()).toEqual([]);

    replica.syncs[0].result.resolve(undefined);
    await settle();
    transport.receive(encodeSyncStep2(Uint8Array.of(0)));
    await settle();
    expect(requests()).toEqual([{ type: 'sync-step-1', stateVector: Uint8Array.of(0) }]);
    expect(errors).toEqual([]);
  });

  it('fails with backpressure once held updates pass maxSyncBytes', async () => {
    const { replica, provider, transport, errors } = await openAsync({ maxSyncBytes: 100 });
    transport.receive(encodeUpdate(Uint8Array.of(1, 2, 3)));
    transport.receive(encodeUpdate(Uint8Array.of(4, 5)));
    await settle();
    expect(errors.map((error) => error.code)).toEqual(['backpressure']);
    expect(provider.status).toBe('disconnected');
    expect(replica.syncs).toEqual([]);
    expect(replica.merges).toEqual([]);
  });

  it('discards a held sync that resolves after a reopen', async () => {
    const { replica, provider, transport } = await openAsync();
    transport.receive(encodeUpdate(Uint8Array.of(1)));
    transport.receive(encodeSyncStep2(Uint8Array.of(0)));
    await settle();
    expect(replica.syncs).toHaveLength(1);
    transport.emit({ type: 'close' });
    transport.emit({ type: 'open' });
    await settle();
    replica.syncs[0].result.resolve(undefined);
    await settle();
    expect(provider.synced).toBe(false);
  });
});

describe('CollaborationProvider limits', () => {
  const replica = new ChunkReplica();
  const transport = new Transport();
  const create = (options: CollaborationProviderOptions) =>
    providers.add(new CollaborationProvider(replica, transport, options));

  it('accepts limits up to their bounds', () => {
    create({ maxFrameBytes: MAX_FRAME_BYTES, syncChunkBytes: MAX_FRAME_BYTES });
    create({ maxPendingBytes: 256 * MiB, maxSyncBytes: 256 * MiB });
    create({ maxPendingBytes: 0, maxSyncBytes: 1, syncChunkBytes: 1 });
    create({ maxFrameBytes: 6 });
  });

  for (const [options, message] of [
    [{ maxFrameBytes: MAX_FRAME_BYTES + 1 }, `maxFrameBytes must be a positive integer no greater than ${MAX_FRAME_BYTES}`],
    [{ maxFrameBytes: 0 }, `maxFrameBytes must be a positive integer no greater than ${MAX_FRAME_BYTES}`],
    [{ maxPendingBytes: 256 * MiB + 1 }, `maxPendingBytes must be a non-negative integer no greater than ${256 * MiB}`],
    [{ syncChunkBytes: 0 }, `syncChunkBytes must be a positive integer no greater than ${16 * MiB}`],
    [{ maxFrameBytes: 4096, syncChunkBytes: 4097 }, 'syncChunkBytes must be a positive integer no greater than 4096'],
    [{ syncChunkBytes: 1.5 }, `syncChunkBytes must be a positive integer no greater than ${16 * MiB}`],
    [{ maxSyncBytes: 0 }, `maxSyncBytes must be a positive integer no greater than ${256 * MiB}`],
    [{ maxSyncBytes: 256 * MiB + 1 }, `maxSyncBytes must be a positive integer no greater than ${256 * MiB}`],
  ] as const) {
    it(`refuses ${JSON.stringify(options)}`, () => {
      expect(() => create(options)).toThrow(new RangeError(message));
    });
  }
});

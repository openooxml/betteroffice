import { afterEach, describe, expect, it } from 'bun:test';
import { createAsyncCollaborationReplica } from './asyncReplica';
import { CollaborationProvider } from './provider';
import {
  decodeAwarenessUpdate,
  decodeMessages,
  encodeAwarenessMessage,
  encodeQueryAwareness,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
} from './protocol';
import type {
  AsyncCollaborationReplica,
  CollaborationError,
  CollaborationPeer,
  CollaborationProviderOptions,
  CollaborationStatusChange,
  CollaborationTransport,
  CollaborationTransportEvent,
  CollaborationUpdateOrigin,
} from './types';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (cause: unknown) => void;
  const promise = new Promise<T>((done, fail) => {
    resolve = done;
    reject = fail;
  });
  return { promise, resolve, reject };
}

class FakeAsyncReplica implements AsyncCollaborationReplica {
  readonly clientId = 1;
  readonly vectors: ReturnType<typeof deferred<Uint8Array>>[] = [];
  readonly diffs: {
    vector: Uint8Array | undefined;
    result: ReturnType<typeof deferred<Uint8Array>>;
  }[] = [];
  readonly merges: {
    update: Uint8Array;
    result: ReturnType<typeof deferred<unknown>>;
  }[] = [];
  private readonly listeners = new Set<
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();

  encodeStateVector(): Promise<Uint8Array> {
    const result = deferred<Uint8Array>();
    this.vectors.push(result);
    return result.promise;
  }

  encodeStateAsUpdate(vector?: Uint8Array): Promise<Uint8Array> {
    const result = deferred<Uint8Array>();
    this.diffs.push({ vector, result });
    return result.promise;
  }

  applyUpdate(update: Uint8Array): Promise<unknown> {
    const result = deferred<unknown>();
    this.merges.push({ update, result });
    return result.promise;
  }

  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  emit(update: Uint8Array, origin: CollaborationUpdateOrigin): void {
    for (const listener of [...this.listeners]) listener(update, origin);
  }
}

class FakeTransport implements CollaborationTransport {
  readonly sent: Uint8Array[] = [];
  disconnectCount = 0;
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();

  connect(): void {}

  disconnect(): void {
    this.disconnectCount += 1;
  }

  send(frame: Uint8Array): boolean {
    this.sent.push(frame.slice());
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

function setup(options: CollaborationProviderOptions = {}) {
  const replica = new FakeAsyncReplica();
  const adapter = createAsyncCollaborationReplica(replica);
  const syncCalls: string[] = [];
  const encodeVector = adapter.encodeStateVector;
  const encodeDiff = adapter.encodeStateAsUpdate;
  adapter.encodeStateVector = () => {
    syncCalls.push('vector');
    return encodeVector();
  };
  adapter.encodeStateAsUpdate = (vector) => {
    syncCalls.push('diff');
    return encodeDiff(vector);
  };
  const transport = new FakeTransport();
  const provider = new CollaborationProvider(adapter, transport, options);
  const errors: CollaborationError[] = [];
  const statuses: CollaborationStatusChange[] = [];
  const peers: (readonly CollaborationPeer[])[] = [];
  provider.onError((error) => errors.push(error));
  provider.onStatus((status) => statuses.push(status));
  provider.onPeers((snapshot) => peers.push(snapshot));
  providers.add(provider);
  return { replica, adapter, syncCalls, transport, provider, errors, statuses, peers };
}

async function settle(): Promise<void> {
  for (let turn = 0; turn < 20; turn++) await Promise.resolve();
}

async function open(value: ReturnType<typeof setup>): Promise<void> {
  value.provider.connect();
  value.transport.emit({ type: 'open' });
  await settle();
  expect(value.replica.vectors).toHaveLength(1);
  value.replica.vectors[0].resolve(Uint8Array.of(7));
  await settle();
  expect(value.errors).toEqual([]);
}

const inference = { clientId: 8, story: 'body', paraId: 'p1', endOffset: 6 };

function announcePeer(transport: FakeTransport): void {
  transport.receive(encodeAwarenessMessage([{
    clientId: 8,
    clock: 3,
    state: {
      user: { name: 'Bright Fox', color: '#137333' },
      cursor: { story: 'body', anchor: Uint8Array.of(1), head: Uint8Array.of(1) },
    },
  }]));
}

describe('CollaborationProvider async replicas', () => {
  it('sends the resolved vector before querying and broadcasting awareness', async () => {
    const value = setup({ user: { name: 'Calm Otter' } });
    value.provider.connect();
    value.transport.emit({ type: 'open' });
    await settle();
    expect(value.transport.sent).toEqual([]);
    expect(value.provider.status).toBe('connected');
    expect(value.provider.synced).toBe(false);

    value.replica.vectors[0].resolve(Uint8Array.of(7, 9));
    await settle();
    const messages = value.transport.sent.flatMap((frame) => decodeMessages(frame));
    expect(messages.map((message) => message.type)).toEqual([
      'sync-step-1', 'query-awareness', 'awareness',
    ]);
    expect(messages[0]).toEqual({ type: 'sync-step-1', stateVector: Uint8Array.of(7, 9) });
    if (messages[2].type !== 'awareness') throw new Error('expected awareness');
    expect(decodeAwarenessUpdate(messages[2].update)[0].state?.user.name).toBe('Calm Otter');
    expect(value.syncCalls).toEqual([]);
    expect(value.errors).toEqual([]);
  });

  it('answers SyncStep1 with the resolved async diff', async () => {
    const value = setup();
    await open(value);
    value.transport.sent.length = 0;
    const frame = encodeSyncStep1(Uint8Array.of(4, 5));
    value.transport.receive(frame);
    frame.fill(0);
    await settle();
    expect(value.replica.diffs[0].vector).toEqual(Uint8Array.of(4, 5));
    expect(value.transport.sent).toEqual([]);

    value.replica.diffs[0].result.resolve(Uint8Array.of(8, 9));
    await settle();
    expect(value.transport.sent.map((bytes) => decodeMessages(bytes))).toEqual([
      [{ type: 'sync-step-2', update: Uint8Array.of(8, 9) }],
    ]);
    expect(value.syncCalls).toEqual([]);
  });

  it('marks synced and queries awareness only after the SyncStep2 merge resolves', async () => {
    const value = setup({ user: { name: 'Calm Otter' } });
    await open(value);
    announcePeer(value.transport);
    value.transport.sent.length = 0;
    value.transport.receive(encodeSyncStep2(Uint8Array.of(2)));
    await settle();
    expect(value.provider.synced).toBe(false);
    expect(value.transport.sent).toEqual([]);

    value.replica.merges[0].result.resolve(inference);
    await settle();
    expect(value.provider.synced).toBe(true);
    expect(value.statuses.at(-1)).toEqual({ status: 'connected', synced: true });
    expect(value.transport.sent.map((frame) => decodeMessages(frame))).toEqual([
      [{ type: 'query-awareness' }],
    ]);
    expect(value.provider.peers[0].inferredCursor).toBeNull();
  });

  for (const operation of ['vector', 'diff', 'merge', 'update'] as const) {
    it(`fails with replica on a rejected ${operation} and stops queued work`, async () => {
      const value = setup();
      if (operation === 'vector') {
        value.provider.connect();
        value.transport.emit({ type: 'open' });
        await settle();
      } else {
        await open(value);
        value.transport.sent.length = 0;
        value.transport.receive(operation === 'diff'
          ? encodeSyncStep1(Uint8Array.of(2))
          : operation === 'merge'
            ? encodeSyncStep2(Uint8Array.of(2))
            : encodeUpdate(Uint8Array.of(2)));
        value.transport.receive(encodeUpdate(Uint8Array.of(3)));
        await settle();
      }
      const failure = new Error('worker failed');
      if (operation === 'vector') value.replica.vectors[0].reject(failure);
      else if (operation === 'diff') value.replica.diffs[0].result.reject(failure);
      else value.replica.merges[0].result.reject(failure);
      await settle();

      expect(value.provider.status).toBe('disconnected');
      expect(value.provider.synced).toBe(false);
      expect(value.statuses.some((status) => status.synced)).toBe(false);
      expect(value.errors.map((error) => error.code)).toEqual(['replica']);
      expect(value.errors[0].cause).toBe(failure);
      expect(value.transport.disconnectCount).toBe(1);
      expect(value.transport.sent).toEqual([]);
      expect(value.replica.merges).toHaveLength(operation === 'merge' || operation === 'update' ? 1 : 0);
    });
  }

  it('orders diffs, sync merges and updates across frames', async () => {
    const value = setup();
    await open(value);
    value.transport.receive(encodeSyncStep1(Uint8Array.of(1)));
    value.transport.receive(encodeSyncStep2(Uint8Array.of(2)));
    value.transport.receive(encodeUpdate(Uint8Array.of(3)));
    value.transport.receive(encodeUpdate(Uint8Array.of(4)));
    await settle();
    expect(value.replica.diffs).toHaveLength(1);
    expect(value.replica.merges).toEqual([]);

    value.replica.diffs[0].result.resolve(Uint8Array.of(8));
    await settle();
    expect(value.replica.merges.map((merge) => merge.update)).toEqual([Uint8Array.of(2)]);
    value.replica.merges[0].result.resolve(undefined);
    await settle();
    expect(value.provider.synced).toBe(true);
    expect(value.replica.merges.map((merge) => merge.update)).toEqual([
      Uint8Array.of(2), Uint8Array.of(3),
    ]);
    value.replica.merges[1].result.resolve(undefined);
    await settle();
    expect(value.replica.merges.map((merge) => merge.update)).toEqual([
      Uint8Array.of(2), Uint8Array.of(3), Uint8Array.of(4),
    ]);
    value.replica.merges[2].result.resolve(undefined);
    await settle();
    expect(value.errors).toEqual([]);
  });

  for (const operation of ['vector', 'diff', 'merge', 'update'] as const) {
    for (const rejected of [false, true]) {
      it(`discards a stale ${operation} ${rejected ? 'rejection' : 'resolution'} after reopen`, async () => {
        const value = setup({ maxPendingBytes: 4 });
        if (operation === 'vector') {
          value.provider.connect();
          value.transport.emit({ type: 'open' });
          await settle();
        } else {
          await open(value);
          value.transport.receive(operation === 'diff'
            ? encodeSyncStep1(Uint8Array.of(2))
            : operation === 'merge'
              ? encodeSyncStep2(Uint8Array.of(2))
              : encodeUpdate(Uint8Array.of(2)));
          await settle();
        }
        value.transport.emit({ type: 'close' });
        value.transport.emit({ type: 'open' });
        await settle();
        value.replica.vectors[1].resolve(Uint8Array.of(9));
        await settle();
        announcePeer(value.transport);
        value.transport.receive(encodeUpdate(Uint8Array.of(3)));
        await settle();
        const newMerge = value.replica.merges.at(-1)!;
        expect(newMerge.update).toEqual(Uint8Array.of(3));
        const frames = value.transport.sent.length;
        const statuses = value.statuses.length;
        const peers = value.peers.length;

        if (operation === 'vector') {
          if (rejected) value.replica.vectors[0].reject(new Error('stale vector'));
          else value.replica.vectors[0].resolve(Uint8Array.of(7));
        } else if (operation === 'diff') {
          if (rejected) value.replica.diffs[0].result.reject(new Error('stale diff'));
          else value.replica.diffs[0].result.resolve(Uint8Array.of(8));
        } else {
          if (rejected) value.replica.merges[0].result.reject(new Error('stale merge'));
          else value.replica.merges[0].result.resolve(inference);
        }
        await settle();
        expect(value.transport.sent).toHaveLength(frames);
        expect(value.statuses).toHaveLength(statuses);
        expect(value.peers).toHaveLength(peers);
        expect(value.provider.synced).toBe(false);
        expect(value.provider.peers[0].inferredCursor).toBeNull();
        expect(value.errors).toEqual([]);

        newMerge.result.resolve(undefined);
        await settle();
        value.transport.receive(encodeUpdate(Uint8Array.of(4)));
        await settle();
        expect(value.replica.merges.at(-1)?.update).toEqual(Uint8Array.of(4));
        expect(value.errors).toEqual([]);
      });
    }
  }

  for (const cleanup of ['disconnect', 'destroy'] as const) {
    for (const operation of ['vector', 'diff', 'merge', 'update'] as const) {
      it(`drops pending ${operation} work on ${cleanup}`, async () => {
        const value = setup();
        if (operation === 'vector') {
          value.provider.connect();
          value.transport.emit({ type: 'open' });
          await settle();
        } else {
          await open(value);
          value.transport.receive(operation === 'diff'
            ? encodeSyncStep1(Uint8Array.of(2))
            : operation === 'merge'
              ? encodeSyncStep2(Uint8Array.of(2))
              : encodeUpdate(Uint8Array.of(2)));
          await settle();
        }
        value.transport.receive(encodeUpdate(Uint8Array.of(3)));
        value.provider[cleanup]();
        const frames = value.transport.sent.length;
        const statuses = value.statuses.length;
        const peers = value.peers.length;
        if (operation === 'vector') value.replica.vectors[0].resolve(Uint8Array.of(7));
        else if (operation === 'diff') value.replica.diffs[0].result.resolve(Uint8Array.of(8));
        else value.replica.merges[0].result.resolve(inference);
        await settle();

        expect(value.transport.sent).toHaveLength(frames);
        expect(value.statuses).toHaveLength(statuses);
        expect(value.peers).toHaveLength(peers);
        expect(value.provider.status).toBe(cleanup === 'destroy' ? 'destroyed' : 'disconnected');
        expect(value.provider.synced).toBe(false);
        expect(value.errors).toEqual([]);
        expect(value.replica.merges).toHaveLength(operation === 'merge' || operation === 'update' ? 1 : 0);
      });
    }
  }

  it('bounds inbound replica bytes and resets the budget after failure', async () => {
    const value = setup({ maxPendingBytes: 4 });
    await open(value);
    value.transport.receive(encodeSyncStep2(Uint8Array.of(2)));
    await settle();
    value.transport.receive(encodeUpdate(Uint8Array.of(3)));
    expect(value.errors.map((error) => error.code)).toEqual(['backpressure']);
    expect(value.errors[0].message).toBe('Pending replica data exceeds the byte limit');
    expect(value.provider.status).toBe('disconnected');
    expect(value.provider.synced).toBe(false);
    expect(value.transport.disconnectCount).toBe(1);

    value.provider.connect();
    value.transport.emit({ type: 'open' });
    await settle();
    value.replica.vectors[1].resolve(Uint8Array.of(9));
    await settle();
    value.transport.receive(encodeSyncStep2(Uint8Array.of(4)));
    await settle();
    expect(value.replica.merges.map((merge) => merge.update)).toEqual([
      Uint8Array.of(2), Uint8Array.of(4),
    ]);
    value.replica.merges[0].result.resolve(inference);
    await settle();
    expect(value.provider.synced).toBe(false);
    value.replica.merges[1].result.resolve(undefined);
    await settle();
    expect(value.provider.synced).toBe(true);
    expect(value.errors).toHaveLength(1);
  });

  it('handles awareness and queries immediately while an update is pending', async () => {
    const value = setup();
    await open(value);
    value.transport.sent.length = 0;
    value.transport.receive(encodeUpdate(Uint8Array.of(2)));
    await settle();
    announcePeer(value.transport);
    value.transport.receive(encodeQueryAwareness());
    expect(value.provider.peers[0].clientId).toBe(8);
    expect(value.transport.sent.map((frame) => decodeMessages(frame))).toEqual([
      [{ type: 'awareness', update: Uint8Array.of(0) }],
    ]);
    expect(value.errors).toEqual([]);
  });

  it('handles auth immediately and cancels a pending merge', async () => {
    const value = setup();
    await open(value);
    value.transport.receive(encodeSyncStep2(Uint8Array.of(2)));
    await settle();
    value.transport.receive(Uint8Array.of(2, 0, 0));
    expect(value.errors.map((error) => error.code)).toEqual(['protocol']);
    expect(value.errors[0].message).toBe('Authentication denied');
    expect(value.provider.status).toBe('disconnected');
    value.replica.merges[0].result.resolve(undefined);
    await settle();
    expect(value.provider.synced).toBe(false);
    expect(value.errors).toHaveLength(1);
  });

  it('uses the resolved update inference to publish the inferred cursor', async () => {
    const value = setup();
    await open(value);
    announcePeer(value.transport);
    const frame = encodeUpdate(Uint8Array.of(9));
    value.transport.receive(frame);
    frame.fill(0);
    await settle();
    expect(value.replica.merges[0].update).toEqual(Uint8Array.of(9));
    expect(value.provider.peers[0].inferredCursor).toBeNull();

    value.replica.merges[0].result.resolve(inference);
    await settle();
    expect(value.peers.at(-1)?.[0]).toMatchObject({ clientId: 8, inferredCursor: inference });
  });

  it('forwards only local updates through the adapter subscription', async () => {
    const value = setup();
    await open(value);
    value.transport.sent.length = 0;
    value.replica.emit(Uint8Array.of(5), 'remote');
    expect(value.transport.sent).toEqual([]);
    value.replica.emit(Uint8Array.of(6), 'local');
    expect(value.transport.sent.map((frame) => decodeMessages(frame))).toEqual([
      [{ type: 'update', update: Uint8Array.of(6) }],
    ]);
  });

  it('keeps synchronous encoding unavailable and never calls it', async () => {
    const value = setup();
    expect(() => value.adapter.encodeStateVector()).toThrow('requires the async adapter');
    expect(() => value.adapter.encodeStateAsUpdate()).toThrow('requires the async adapter');
    value.syncCalls.length = 0;
    await open(value);
    value.transport.receive(encodeSyncStep1(Uint8Array.of(2)));
    await settle();
    value.replica.diffs[0].result.resolve(Uint8Array.of(8));
    await settle();
    expect(value.syncCalls).toEqual([]);
    expect(value.errors).toEqual([]);
  });
});

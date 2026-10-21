import { expect, spyOn, test } from 'bun:test';
import type { PipeOp, PipeOutput } from '../../../../shared/collaboration-pipe/protocol';
import type { PipeArrival } from '../../../../shared/collaboration-pipe/worker';
import { attachCollaborationPipe, createAsyncCollaborationReplica } from './asyncReplica';
import { validAwarenessCursor } from './awareness';
import { CollaborationProvider, providerArrival } from './provider';
import * as protocol from './protocol';
import type { AwarenessCursor, AwarenessPeer, CollaborationError, CollaborationStatusChange, CollaborationTransportEvent } from './types';

async function settle() { for (let turn = 0; turn < 12; turn++) await Promise.resolve(); }

for (const fails of [false, true]) {
  test(`piped reconnect waits for asynchronous teardown ${fails ? 'and cancels when it rejects' : 'to finish'}`, async () => {
    const replica = createAsyncCollaborationReplica({ clientId: 73,
      encodeStateVector: async () => Uint8Array.of(0), encodeStateAsUpdate: async () => Uint8Array.of(0, 0),
      applyUpdate: async () => null, onUpdate: () => () => {},
    });
    attachCollaborationPipe(replica, { open: () => ({ post: async () => {}, close() {} }) });
    let finish!: () => void;
    let reject!: (cause: unknown) => void;
    const cleanup = new Promise<void>((resolve, fail) => { finish = resolve; reject = fail; });
    let connects = 0;
    let disconnects = 0;
    let subscriptions = 0;
    const provider = new CollaborationProvider(replica, {
      connect() { connects += 1; }, disconnect() { disconnects += 1; return cleanup; },
      send: () => true, onEvent() { subscriptions += 1; return () => {}; },
    });
    const errors: CollaborationError[] = [];
    provider.onError((error) => errors.push(error));
    try {
      provider.connect();
      provider.disconnect();
      provider.connect();
      expect(provider.status).toBe('connecting');
      expect(connects).toBe(1);
      expect(disconnects).toBe(1);
      expect(subscriptions).toBe(1);
      if (fails) reject(new Error('Still connected'));
      else finish();
      await settle();
      expect(connects).toBe(fails ? 1 : 2);
      expect(disconnects).toBe(1);
      expect(subscriptions).toBe(fails ? 1 : 2);
      expect(provider.status).toBe(fails ? 'disconnected' : 'connecting');
      expect(errors.map(({ code, message }) => ({ code, message }))).toEqual(fails
        ? [{ code: 'transport', message: 'Transport disconnect failed: Still connected' }] : []);
    } finally { provider.destroy(); await settle(); }
  });
}

for (const fails of [false, true]) {
  test(fails ? 'refused initial sync never reports synced' : 'initial sync waits for its arrival to merge', () => {
    let receive!: (event: CollaborationTransportEvent) => void;
    const pending: PipeArrival[] = [];
    const provider = new CollaborationProvider({ clientId: 74,
      encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0),
      applyUpdate() { pending.push(providerArrival(provider)); }, onUpdate: () => () => {},
    }, { connect() { receive({ type: 'open' }); }, disconnect() {}, send: () => true,
      onEvent(listener) { receive = listener; return () => {}; } });
    const statuses: CollaborationStatusChange[] = [];
    const errors: CollaborationError[] = [];
    provider.onStatus((status) => statuses.push(status));
    provider.onError((error) => errors.push(error));
    try {
      provider.connect();
      receive({ type: 'message', data: protocol.encodeUpdate(Uint8Array.of(0, 0)) });
      pending[0].settle(null);
      expect(provider.synced).toBe(false);
      receive({ type: 'message', data: protocol.encodeSyncStep2(Uint8Array.of(0, 0)) });
      expect(provider.synced).toBe(false);
      expect(statuses).toEqual([{ status: 'connecting', synced: false }, { status: 'connected', synced: false }]);
      if (fails) pending[1].fail(new Error('Merge refused'));
      else pending[1].settle(null);
      pending[1].settle(null);
      expect(statuses).toEqual([{ status: 'connecting', synced: false }, { status: 'connected', synced: false },
        fails ? { status: 'disconnected', synced: false } : { status: 'connected', synced: true }]);
      expect(errors.map((error) => error.code)).toEqual(fails ? ['replica'] : []);
    } finally { provider.destroy(); }
  });
}

test('the main pipe facade forwards bytes without decoding or using the async replica', async () => {
  const decode = spyOn(protocol, 'decodeMessages');
  const unavailable = async (): Promise<never> => { throw new Error('Main must not read or apply the replica'); };
  const replica = createAsyncCollaborationReplica({ clientId: 71, encodeStateVector: unavailable,
    encodeStateAsUpdate: unavailable, applyUpdate: unavailable,
    onUpdate: () => { throw new Error('Main must not observe the replica'); } });
  const posts: PipeOp<AwarenessCursor>[][] = [];
  let output!: (out: readonly PipeOutput<AwarenessPeer>[]) => void;
  attachCollaborationPipe(replica, { open(receive) {
    output = receive;
    return { async post(ops) { posts.push(ops); }, close() {} };
  } });
  let receive!: (event: CollaborationTransportEvent) => void;
  const provider = new CollaborationProvider(replica, {
    connect() { receive({ type: 'open' }); }, disconnect() {}, send: () => true,
    onEvent(listener) { receive = listener; return () => {}; },
  });
  const errors: CollaborationError[] = [];
  provider.onError((error) => errors.push(error));
  try {
    provider.connect();
    await settle();
    const bytes = Uint8Array.of(0xff, 0xff, 0xff);
    for (let index = 0; index < 20; index++) receive({ type: 'message', data: bytes });
    bytes.fill(0);
    await settle();
    expect(posts).toHaveLength(2);
    const events = posts[1].filter((op) => op.op === 'event');
    expect(events).toHaveLength(20);
    for (const op of events) {
      expect(op.event.type).toBe('message');
      if (op.event.type === 'message') expect([...new Uint8Array(op.event.data)]).toEqual([0xff, 0xff, 0xff]);
    }
    expect(decode).not.toHaveBeenCalled();
    expect(errors).toEqual([]);
    const through = posts[1][posts[1].length - 1].seq;
    output([{ do: 'status', through, status: 'connected', synced: true }]);
    expect(provider.status).toBe('connected');
    expect(provider.synced).toBe(true);
    const cursor = { sheet: 'Sheet', anchor: { row: 2, col: 3 }, head: { row: 4, col: 5 } };
    provider.setCursor(cursor);
    cursor.head.row = 99;
    await settle();
    expect(posts[2]).toMatchObject([{ op: 'cursor', cursor: { head: { row: 4 } } }]);
    provider.setCursor({ ...cursor, sheet: '' });
    expect(errors[0]).toMatchObject({ code: 'protocol' });
    await settle();
    expect(posts[3]).toMatchObject([{ op: 'cursor', cursor: null }]);
  } finally { provider.destroy(); await settle(); decode.mockRestore(); }
});

test('piped peers keep their wall-clock timestamps whatever the worker time origin', async () => {
  const replica = createAsyncCollaborationReplica({ clientId: 75,
    encodeStateVector: async () => Uint8Array.of(0), encodeStateAsUpdate: async () => Uint8Array.of(0, 0),
    applyUpdate: async () => null, onUpdate: () => () => {},
  });
  let output!: (out: readonly PipeOutput<AwarenessPeer>[]) => void;
  attachCollaborationPipe(replica, { open(receive) {
    output = receive;
    return { async post() {}, close() {} };
  } });
  const provider = new CollaborationProvider(replica, {
    connect() {}, disconnect() {}, send: () => true, onEvent: () => () => {},
  });
  const received: Array<readonly AwarenessPeer[]> = [];
  provider.onAwareness(peers => received.push(peers));
  const peer: AwarenessPeer = { clientId: 8, clock: 1, user: { name: 'B', color: '#000' },
    cursor: { sheet: 'Sheet', anchor: { row: 2, col: 3 }, head: { row: 4, col: 5 } },
    lastSeen: 100, cursorMovedAt: 80 };
  try {
    provider.connect();
    for (const offset of [20_000, -20_000]) {
      output([{ do: 'peers', through: 2, timeOrigin: performance.timeOrigin + offset,
        peers: [peer, { ...peer, clientId: 9, cursor: null, cursorMovedAt: 0 }] }]);
      expect(provider.peers).toEqual([
        peer,
        { ...peer, clientId: 9, cursor: null, cursorMovedAt: 0 },
      ]);
      expect(received[received.length - 1]).toBe(provider.peers);
      expect(peer).toMatchObject({ lastSeen: 100, cursorMovedAt: 80 });
    }
  } finally { provider.destroy(); await settle(); }
});

test('deferred arrivals fail once and stale failures cannot disconnect a reopened sync provider', () => {
  let receive!: (event: CollaborationTransportEvent) => void;
  const provider = new CollaborationProvider({ clientId: 72,
    encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0),
    applyUpdate: () => null, onUpdate: () => () => {},
  }, { connect() { receive({ type: 'open' }); }, disconnect() {}, send: () => true,
    onEvent(listener) { receive = listener; return () => {}; } });
  const errors: CollaborationError[] = [];
  provider.onError((error) => errors.push(error));
  try {
    provider.connect();
    const stale = providerArrival(provider);
    provider.disconnect();
    provider.connect();
    stale.fail(new Error('Old merge failed'));
    expect(errors).toEqual([]);
    expect(provider.status).toBe('connected');
    const previousOpen = providerArrival(provider);
    receive({ type: 'close' });
    receive({ type: 'open' });
    previousOpen.fail(new Error('Previous open failed'));
    expect(errors).toEqual([]);
    expect(provider.status).toBe('connected');
    receive({ type: 'message', data: protocol.encodeSyncStep2(Uint8Array.of(0, 0)) });
    expect(provider.synced).toBe(true);
    const settled = providerArrival(provider);
    settled.settle(null);
    settled.fail(new Error('Already settled'));
    expect(errors).toEqual([]);
    const arrival = providerArrival(provider);
    arrival.fail(new Error('Merge refused'));
    arrival.fail(new Error('Already failed'));
    expect(errors).toHaveLength(1);
    expect(errors[0]).toMatchObject({ code: 'replica', message: 'Failed to apply remote update: Merge refused' });
    expect(provider.status).toBe('disconnected');
  } finally { provider.destroy(); }
});

test('pipe cursor validation retains XLSX encodable cursor rules', () => {
  const cursor = { sheet: 'Sheet', anchor: { row: 0, col: 0 }, head: { row: 1_048_576, col: 16_384 } };
  expect(validAwarenessCursor(cursor)).toBe(true);
  for (const invalid of [null, {}, { ...cursor, sheet: '' }, { ...cursor, sheet: 's'.repeat(257) },
    { ...cursor, anchor: { row: -1, col: 0 } }, { ...cursor, head: { row: 0.5, col: 0 } },
    { ...cursor, head: { row: 0, col: Number.POSITIVE_INFINITY } }]) {
    expect(validAwarenessCursor(invalid)).toBe(false);
  }
});

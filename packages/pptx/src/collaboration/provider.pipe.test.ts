import { expect, spyOn, test } from 'bun:test';
import type { PipeOp, PipeOutput } from '../../../../shared/collaboration-pipe/protocol';
import { attachCollaborationPipe } from './asyncReplica';
import { validPresenceCursor } from './presence';
import { CollaborationProvider, providerArrival } from './provider';
import * as protocol from './protocol';
import type { PptxPresenceCursor, PptxPresencePeer, CollaborationError, CollaborationTransportEvent } from './types';

async function settle() { for (let turn = 0; turn < 12; turn++) await Promise.resolve(); }

test('identical normalized cursor republishes stop issuing pipe commands', async () => {
  const unavailable = async (): Promise<never> => { throw new Error('Unexpected replica access'); };
  const replica = { clientId: 71, encodeStateVector: unavailable, encodeStateAsUpdate: unavailable,
    applyUpdate: unavailable, onUpdate: () => () => {} };
  const posts: PipeOp<PptxPresenceCursor>[][] = [];
  let provider!: CollaborationProvider;
  attachCollaborationPipe(replica, { open() {
    return { async post(ops) {
      posts.push(ops);
      const cursor = ops.find((op) => op.op === 'cursor');
      if (posts.length < 8 && cursor?.op === 'cursor') provider.setCursor(cursor.cursor && { ...cursor.cursor });
    }, close() {} };
  } });
  provider = new CollaborationProvider(replica, {
    connect() {}, disconnect() {}, send: () => true, onEvent: () => () => {},
  });
  try {
    await settle();
    const cursor = { slideId: 'slide', shapeId: 'shape' };
    provider.setCursor(cursor);
    provider.setCursor({ ...cursor });
    cursor.shapeId = 'changed';
    await settle();
    expect(posts).toHaveLength(2);
    expect(posts[1]).toHaveLength(1);
    expect(posts[1]).toMatchObject([{ op: 'cursor', cursor: { slideId: 'slide', shapeId: 'shape' } }]);
    provider.setCursor({ slideId: 'slide', shapeId: 'shape' });
    await settle();
    expect(posts).toHaveLength(2);
    provider.setCursor(null);
    await settle();
    expect(posts).toHaveLength(3);
    expect(posts[2]).toMatchObject([{ op: 'cursor', cursor: null }]);
  } finally { provider.destroy(); await settle(); }
});

for (const result of ['settle', 'fail', 'stale'] as const) test(`deferred SyncStep2 ${result} reports sync only after a successful merge`, () => {
  let receive!: (event: CollaborationTransportEvent) => void;
  let provider!: CollaborationProvider;
  const arrivals: ReturnType<typeof providerArrival>[] = [];
  provider = new CollaborationProvider({ clientId: 72,
    encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0),
    applyUpdate() { arrivals.push(providerArrival(provider)); }, onUpdate: () => () => {},
  }, { connect() { receive({ type: 'open' }); }, disconnect() {}, send: () => true,
    onEvent(listener) { receive = listener; return () => {}; } });
  const synced: boolean[] = [];
  const errors: CollaborationError[] = [];
  provider.onStatus((change) => synced.push(change.synced));
  provider.onError((error) => errors.push(error));
  try {
    provider.connect();
    receive({ type: 'message', data: protocol.encodeSyncStep2(Uint8Array.of(0, 0)) });
    receive({ type: 'message', data: protocol.encodeUpdate(Uint8Array.of(0, 0)) });
    expect(arrivals).toHaveLength(2);
    expect(provider.synced).toBe(false);
    expect(synced).not.toContain(true);
    arrivals[1]!.settle(null);
    expect(provider.synced).toBe(false);
    if (result === 'settle') {
      arrivals[0]!.settle(null);
      expect(provider.synced).toBe(true);
      expect(synced.filter(Boolean)).toHaveLength(1);
      expect(errors).toEqual([]);
    } else if (result === 'fail') {
      arrivals[0]!.fail(new Error('Merge refused'));
      arrivals[0]!.settle(null);
      expect(provider.status).toBe('disconnected');
      expect(provider.synced).toBe(false);
      expect(synced).not.toContain(true);
      expect(errors).toMatchObject([{ code: 'replica' }]);
    } else {
      receive({ type: 'close' });
      receive({ type: 'open' });
      arrivals[0]!.settle(null);
      expect(provider.status).toBe('connected');
      expect(provider.synced).toBe(false);
      expect(synced).not.toContain(true);
      receive({ type: 'message', data: protocol.encodeSyncStep2(Uint8Array.of(0, 0)) });
      expect(arrivals).toHaveLength(3);
      expect(provider.synced).toBe(false);
      arrivals[2]!.settle(null);
      expect(provider.synced).toBe(true);
      expect(errors).toEqual([]);
    }
  } finally { provider.destroy(); }
});

test('the main pipe facade forwards bytes without decoding or using the async replica', async () => {
  const decode = spyOn(protocol, 'decodeMessages');
  const unavailable = async (): Promise<never> => { throw new Error('Main must not read or apply the replica'); };
  const replica = { clientId: 71, encodeStateVector: unavailable,
    encodeStateAsUpdate: unavailable, applyUpdate: unavailable,
    onUpdate: () => { throw new Error('Main must not observe the replica'); } };
  const posts: PipeOp<PptxPresenceCursor>[][] = [];
  let output!: (out: readonly PipeOutput<PptxPresencePeer>[]) => void;
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
    const cursor = { slideId: 'slide1', shapeId: 'shape4' };
    provider.setCursor(cursor);
    cursor.shapeId = 'shape99';
    await settle();
    expect(posts[2]).toMatchObject([{ op: 'cursor', cursor: { shapeId: 'shape4' } }]);
    provider.setCursor({ ...cursor, slideId: 's'.repeat(1025) });
    expect(errors[0]).toMatchObject({ code: 'protocol' });
    await settle();
    expect(posts[3]).toMatchObject([{ op: 'cursor', cursor: null }]);
  } finally { provider.destroy(); await settle(); decode.mockRestore(); }
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

test('pipe cursor validation retains PPTX encodable cursor rules', () => {
  expect(validPresenceCursor({ slideId: 'slide' }, 4096)).toBe(true);
  expect(validPresenceCursor({ slideId: 'slide', shapeId: 'shape' }, 4096)).toBe(true);
  for (const invalid of [null, {}, { slideId: '' }, { slideId: 's'.repeat(1025) },
    { slideId: 'slide', shapeId: '' }, { slideId: 'slide', shapeId: 1 },
    { slideId: 'slide', shapeId: 's'.repeat(1025) }]) {
    expect(validPresenceCursor(invalid, 4096)).toBe(false);
  }
  expect(validPresenceCursor({ slideId: 'slide' }, 1)).toBe(false);
});

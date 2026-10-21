import { expect, test } from 'bun:test';
import { CollaborationProvider } from './provider';
import { encodeAwarenessUpdate, encodeSyncStep1, encodeSyncStep2, encodeUpdate } from './protocol';
import type { AsyncCollaborationReplica, CollaborationError, CollaborationTransport, CollaborationTransportEvent } from './types';

function deferred() {
  let resolve!: () => void;
  let reject!: (cause: unknown) => void;
  const promise = new Promise<void>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function setup(options: { maxFrameBytes?: number; maxPendingBytes?: number; chunks?: boolean } = {}) {
  const listeners = new Set<(event: CollaborationTransportEvent) => void>();
  const errors: CollaborationError[] = [];
  const calls: number[][] = [];
  const gates: ReturnType<typeof deferred>[] = [];
  const hold = () => { const gate = deferred(); gates.push(gate); return gate.promise; };
  const replica: AsyncCollaborationReplica & { applyUpdateChunks?(updates: readonly Uint8Array[]): Promise<unknown> } = {
    clientId: 1,
    encodeStateVector: async () => Uint8Array.of(1),
    encodeStateAsUpdate: async () => Uint8Array.of(2),
    async applyUpdate(update) { calls.push([update[0]!]); await hold(); },
    onUpdate: () => () => {},
    ...(options.chunks !== false && {
      async applyUpdateChunks(updates: readonly Uint8Array[]) { calls.push(updates.map((update) => update[0]!)); await hold(); },
    }),
  };
  const transport: CollaborationTransport = { connect() {}, disconnect() {}, send: () => true,
    onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; } };
  const provider = new CollaborationProvider(replica, transport, { maxFrameBytes: options.maxFrameBytes, maxPendingBytes: options.maxPendingBytes });
  provider.onError((error) => { errors.push(error); });
  const emit = (data: Uint8Array) => { for (const listener of [...listeners]) listener({ type: 'message', data }); };
  provider.connect();
  for (const listener of [...listeners]) listener({ type: 'open' });
  const release = async () => { gates.shift()?.resolve(); await settle(); };
  return { provider, calls, emit, errors, release, gates };
}

async function settle() { for (let index = 0; index < 20; index++) await Promise.resolve(); }

test('updates queued behind a busy replica apply as one chunked call in arrival order', async () => {
  const value = setup();
  try {
    await settle();
    for (const byte of [3, 4, 5, 6]) value.emit(encodeUpdate(Uint8Array.of(byte)));
    await settle();
    expect(value.calls).toEqual([[3]]);
    await value.release();
    expect(value.calls).toEqual([[3], [4, 5, 6]]);
    await value.release();
    expect(value.gates).toEqual([]);
    expect(value.errors).toEqual([]);
  } finally { value.provider.destroy(); }
});

test('queued updates never merge across sync steps', async () => {
  const value = setup();
  try {
    await settle();
    value.emit(encodeUpdate(Uint8Array.of(3)));
    value.emit(encodeUpdate(Uint8Array.of(4)));
    value.emit(encodeUpdate(Uint8Array.of(5)));
    value.emit(encodeSyncStep2(Uint8Array.of(6)));
    value.emit(encodeUpdate(Uint8Array.of(7)));
    value.emit(encodeSyncStep1(Uint8Array.of(0)));
    value.emit(encodeUpdate(Uint8Array.of(8)));
    value.emit(encodeUpdate(Uint8Array.of(9)));
    await settle();
    for (let index = 0; index < 5; index++) await value.release();
    expect(value.calls).toEqual([[3], [4, 5], [6], [7], [8, 9]]);
    expect(value.provider.synced).toBe(true);
    expect(value.errors).toEqual([]);
  } finally { value.provider.destroy(); }
});

test('a chunked apply stays within the frame byte limit', async () => {
  const update = (byte: number) => new Uint8Array(40).fill(byte);
  const value = setup({ maxFrameBytes: 100 });
  try {
    await settle();
    for (const byte of [3, 4, 5, 6, 7]) value.emit(encodeUpdate(update(byte)));
    await settle();
    for (let index = 0; index < 3; index++) await value.release();
    expect(value.calls).toEqual([[3], [4, 5], [6, 7]]);
    expect(value.errors).toEqual([]);
  } finally { value.provider.destroy(); }
});

test('a chunked apply stays within the worker update limit when the frame limits are larger', async () => {
  const update = (byte: number) => new Uint8Array(6 * 1024 * 1024).fill(byte);
  const value = setup({ maxFrameBytes: 64 * 1024 * 1024, maxPendingBytes: 64 * 1024 * 1024 });
  try {
    await settle();
    for (const byte of [3, 4, 5, 6, 7]) value.emit(encodeUpdate(update(byte), 64 * 1024 * 1024));
    await settle();
    for (let index = 0; index < 3; index++) await value.release();
    expect(value.calls).toEqual([[3], [4, 5], [6, 7]]);
    expect(value.errors).toEqual([]);
  } finally { value.provider.destroy(); }
});

test('awareness queued between updates is not held behind later updates', async () => {
  const value = setup();
  const seen: number[][] = [];
  value.provider.onPresence(() => seen.push([value.calls.length]));
  try {
    await settle();
    value.emit(encodeUpdate(Uint8Array.of(3)));
    value.emit(encodeUpdate(Uint8Array.of(4)));
    value.emit(encodeUpdate(Uint8Array.of(5)));
    value.emit(encodeAwarenessUpdate([{ clientId: 2, clock: 1, state: { clientId: 2, clock: 1,
      user: { name: 'Calm Fox', color: '#2E7D32' }, cursor: { slideId: 'slide-a' } } }]));
    value.emit(encodeUpdate(Uint8Array.of(6)));
    value.emit(encodeUpdate(Uint8Array.of(7)));
    await settle();
    await value.release();
    expect(value.calls).toEqual([[3], [4, 5]]);
    await value.release();
    expect(value.provider.peers.map((peer) => peer.state.clientId)).toEqual([2]);
    expect(value.calls).toEqual([[3], [4, 5], [6, 7]]);
    expect(seen[seen.length - 1]).toEqual([2]);
    await value.release();
    expect(value.errors).toEqual([]);
  } finally { value.provider.destroy(); }
});

test('a replica without chunked applies keeps one apply per update', async () => {
  const value = setup({ chunks: false });
  try {
    await settle();
    for (const byte of [3, 4, 5]) value.emit(encodeUpdate(Uint8Array.of(byte)));
    await settle();
    for (let index = 0; index < 3; index++) await value.release();
    expect(value.calls).toEqual([[3], [4], [5]]);
    expect(value.errors).toEqual([]);
  } finally { value.provider.destroy(); }
});

test('a failed chunked apply disconnects without applying later updates', async () => {
  const value = setup();
  try {
    await settle();
    for (const byte of [3, 4, 5]) value.emit(encodeUpdate(Uint8Array.of(byte)));
    await settle();
    await value.release();
    value.emit(encodeSyncStep2(Uint8Array.of(6)));
    value.gates.shift()!.reject(new Error('refused'));
    await settle();
    expect(value.calls).toEqual([[3], [4, 5]]);
    expect(value.errors.map((error) => error.code)).toEqual(['replica']);
    expect(value.provider.status).toBe('disconnected');
  } finally { value.provider.destroy(); }
});

import { expect, test } from 'bun:test';
import { createAsyncCollaborationReplica } from './asyncReplica';
import { CollaborationProvider } from './provider';
import { decodeMessages, encodeSyncStep1, encodeSyncStep2, encodeUpdate } from './protocol';
import type { CollaborationTransport, CollaborationTransportEvent } from './types';
import type { WorkbookSession } from '../session/client';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from '../session/testHelpers';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

test('async handshakes and merges wait for acknowledgements and ignore closed connections', async () => {
  const vector = deferred<Uint8Array>();
  const diff = deferred<Uint8Array>();
  const merge = deferred<void>();
  const applied: number[] = [];
  const frames: Uint8Array[] = [];
  let receive!: (event: CollaborationTransportEvent) => void;
  const replica = createAsyncCollaborationReplica({ clientId: 101,
    encodeStateVector: () => vector.promise,
    encodeStateAsUpdate: () => diff.promise,
    applyUpdate: async (update) => { applied.push(update[0]); await merge.promise; },
    onUpdate: () => () => {},
  });
  const provider = new CollaborationProvider(replica, {
    connect() { receive({ type: 'open' }); }, disconnect() {},
    send(frame) { frames.push(frame); return true; },
    onEvent(listener) { receive = listener; return () => {}; },
  });
  const settle = async () => { for (let turn = 0; turn < 12; turn++) await Promise.resolve(); };
  try {
    provider.connect();
    expect(frames).toHaveLength(0);
    vector.resolve(Uint8Array.of(1));
    await settle();
    expect(decodeMessages(frames[0])[0].type).toBe('sync-step-1');
    receive({ type: 'message', data: encodeSyncStep2(Uint8Array.of(2)) });
    receive({ type: 'message', data: encodeUpdate(Uint8Array.of(3)) });
    await settle();
    expect(provider.synced).toBe(false);
    expect(applied).toEqual([2]);
    merge.resolve();
    await settle();
    expect(provider.synced).toBe(true);
    expect(applied).toEqual([2, 3]);
    receive({ type: 'message', data: encodeSyncStep1(Uint8Array.of(4)) });
    await settle();
    receive({ type: 'close' });
    receive({ type: 'open' });
    await settle();
    const count = frames.length;
    diff.resolve(Uint8Array.of(5));
    await settle();
    expect(frames).toHaveLength(count);
    expect(provider.synced).toBe(false);
  } finally { provider.destroy(); }
});

test('async replica queues are bounded and stale merge acknowledgements cannot sync a reopened connection', async () => {
  const merge = deferred<void>();
  let receive!: (event: CollaborationTransportEvent) => void;
  let applied = 0;
  const errors: string[] = [];
  const replica = createAsyncCollaborationReplica({ clientId: 101,
    encodeStateVector: async () => Uint8Array.of(1),
    encodeStateAsUpdate: async () => Uint8Array.of(2),
    applyUpdate: () => { applied++; return merge.promise; }, onUpdate: () => () => {},
  });
  const provider = new CollaborationProvider(replica, {
    connect() { receive({ type: 'open' }); }, disconnect() {}, send: () => true,
    onEvent(listener) { receive = listener; return () => {}; },
  }, { maxPendingBytes: 4 });
  provider.onError((error) => errors.push(error.code));
  const settle = async () => { for (let turn = 0; turn < 12; turn++) await Promise.resolve(); };
  try {
    provider.connect();
    await settle();
    receive({ type: 'message', data: encodeSyncStep2(Uint8Array.of(3)) });
    await settle();
    expect(applied).toBe(1);
    expect(provider.synced).toBe(false);
    receive({ type: 'message', data: encodeUpdate(Uint8Array.of(4)) });
    expect(errors).toEqual(['backpressure']);
    expect(provider.status).toBe('disconnected');
    provider.connect();
    await settle();
    merge.resolve();
    await settle();
    expect(provider.synced).toBe(false);
    expect(applied).toBe(1);
  } finally { provider.destroy(); }
});

test('async remote updates that arrive during a busy merge land in one chunked apply, in order', async () => {
  const merge = deferred<void>();
  const calls: number[][] = [];
  const frames: Uint8Array[] = [];
  let receive!: (event: CollaborationTransportEvent) => void;
  const replica = createAsyncCollaborationReplica({ clientId: 101,
    encodeStateVector: async () => Uint8Array.of(1),
    encodeStateAsUpdate: async () => Uint8Array.of(9),
    applyUpdate: async (update) => { calls.push([update[0]]); if (calls.length === 1) await merge.promise; },
    applyUpdateChunks: async (updates) => { calls.push(updates.map((update) => update[0])); },
    onUpdate: () => () => {},
  });
  const provider = new CollaborationProvider(replica, {
    connect() { receive({ type: 'open' }); }, disconnect() {},
    send(frame) { frames.push(frame); return true; },
    onEvent(listener) { receive = listener; return () => {}; },
  }, { maxPendingBytes: 64 });
  const settle = async () => { for (let turn = 0; turn < 4; turn++) await new Promise((resolve) => setTimeout(resolve, 0)); };
  try {
    provider.connect();
    await settle();
    receive({ type: 'message', data: encodeSyncStep2(Uint8Array.of(2)) });
    for (const value of [3, 4, 5, 6]) receive({ type: 'message', data: encodeUpdate(Uint8Array.of(value)) });
    receive({ type: 'message', data: encodeSyncStep1(Uint8Array.of(0)) });
    for (const value of [7, 8]) receive({ type: 'message', data: encodeUpdate(Uint8Array.of(value)) });
    await settle();
    expect(calls).toEqual([[2]]);
    merge.resolve();
    await settle();
    expect(calls).toEqual([[2], [3, 4, 5, 6], [7, 8]]);
    expect(provider.synced).toBe(true);
    expect(frames.flatMap((frame) => decodeMessages(frame)).filter((message) => message.type === 'sync-step-2')).toHaveLength(1);
    for (let round = 0; round < 8; round++) {
      receive({ type: 'message', data: encodeUpdate(Uint8Array.of(10 + round)) });
      await settle();
    }
    expect(calls.slice(3)).toEqual(Array.from({ length: 8 }, (_, round) => [10 + round]));
    expect(provider.status).toBe('connected');
  } finally { provider.destroy(); }
});

test('a refused chunked merge fails the connection once', async () => {
  const merge = deferred<void>();
  const errors: string[] = [];
  let chunks = 0;
  let receive!: (event: CollaborationTransportEvent) => void;
  const replica = createAsyncCollaborationReplica({ clientId: 101,
    encodeStateVector: async () => Uint8Array.of(1),
    encodeStateAsUpdate: async () => Uint8Array.of(9),
    applyUpdate: () => merge.promise,
    applyUpdateChunks: async () => { chunks++; throw new Error('invalid-update'); },
    onUpdate: () => () => {},
  });
  const provider = new CollaborationProvider(replica, {
    connect() { receive({ type: 'open' }); }, disconnect() {}, send: () => true,
    onEvent(listener) { receive = listener; return () => {}; },
  });
  provider.onError((error) => errors.push(error.code));
  const settle = async () => { for (let turn = 0; turn < 12; turn++) await Promise.resolve(); };
  try {
    provider.connect();
    await settle();
    receive({ type: 'message', data: encodeUpdate(Uint8Array.of(2)) });
    for (const value of [3, 4]) receive({ type: 'message', data: encodeUpdate(Uint8Array.of(value)) });
    await settle();
    merge.resolve();
    await settle();
    expect(chunks).toBe(1);
    expect(errors).toEqual(['replica']);
    expect(provider.status).toBe('disconnected');
  } finally { provider.destroy(); }
});

test('worker replicas exchange a fresh handshake and converge after disconnected edits', async () => {
  const { fixture: bytes } = await loadWorkbookSessionFixtures();
  const left = await createTestWorkbookSession(bytes, undefined, { collaborative: true, clientId: 101 });
  let right: WorkbookSession | undefined;
  const providers: CollaborationProvider[] = [];
  const listeners = [new Set<(event: CollaborationTransportEvent) => void>(),
    new Set<(event: CollaborationTransportEvent) => void>()];
  const frames: Uint8Array[][] = [[], []];
  const opened = [false, false];
  const connection = (index: number, open: boolean) => {
    opened[index] = open;
    for (const listener of listeners[index]) listener({ type: open ? 'open' : 'close' });
  };
  const transports: CollaborationTransport[] = listeners.map((local, index) => ({
    connect() { connection(index, true); }, disconnect() { connection(index, false); },
    send(data) {
      if (!opened[index] || !opened[1 - index]) return false;
      frames[index].push(data.slice());
      for (const listener of listeners[1 - index]) listener({ type: 'message', data: data.slice() });
      return true;
    },
    onEvent(listener) { local.add(listener); return () => { local.delete(listener); }; },
  }));
  const whenSynced = (provider: CollaborationProvider) => new Promise<void>((resolve) => {
    const off = provider.onStatus(({ synced }) => { if (synced) { off(); resolve(); } });
  });
  const edit = async (session: WorkbookSession, col: number, input: string) => {
    const command = session.editSession!.submit('editCell', { payload: { args: [0, 2, col, input] } });
    expect((await command.ack).status).toBe('applied');
  };
  const count = (index: number, type: string) => frames[index].flatMap((frame) => decodeMessages(frame))
    .filter((message) => message.type === type).length;
  const quiet = async () => {
    for (let idle = 0, seen = -1; idle < 64;) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      const moved = frames[0].length + frames[1].length;
      idle = moved === seen ? idle + 1 : 0;
      seen = moved;
    }
  };
  try {
    right = await createTestWorkbookSession(bytes, undefined, { collaborative: true, clientId: 202 });
    providers.push(new CollaborationProvider(left.replica!, transports[0]),
      new CollaborationProvider(right.replica!, transports[1]));
    let synced = Promise.all(providers.map(whenSynced));
    for (const provider of providers) provider.connect();
    await synced;
    await Promise.all([left.call.version(), right.call.version()]);
    await quiet();
    const handshakes = frames.map((_, index) => [count(index, 'sync-step-1'), count(index, 'sync-step-2')]);
    expect(handshakes).toEqual([[2, 2], [2, 2]]);
    connection(0, false);
    connection(1, false);
    expect(providers.every((provider) => provider.status === 'disconnected' && !provider.synced)).toBe(true);
    const sent = frames.map((messages) => messages.length);
    await Promise.all([edit(left, 1, 'left offline'), edit(right, 2, 'right offline')]);
    const offline = await Promise.all([left.call.cellInputs(0, 'B3:C3'), right.call.cellInputs(0, 'B3:C3')]);
    expect(offline[0].cells).not.toEqual(offline[1].cells);
    expect(frames.map((messages) => messages.length)).toEqual(sent);
    synced = Promise.all(providers.map(whenSynced));
    connection(0, true);
    connection(1, true);
    await synced;
    await Promise.all([left.call.version(), right.call.version()]);
    await quiet();
    expect(providers.every((provider) => provider.synced)).toBe(true);
    expect(frames.map((_, index) => [count(index, 'sync-step-1'), count(index, 'sync-step-2')]))
      .toEqual(handshakes.map(([step1, step2]) => [step1 + 2, step2 + 2]));
    const cells = await Promise.all([left.call.cellInputs(0, 'B3:C3'), right.call.cellInputs(0, 'B3:C3')]);
    expect(cells[0].cells).toEqual(cells[1].cells);
    expect(cells[0].cells[0].map((cell) => cell.input)).toEqual(['left offline', 'right offline']);
  } finally {
    for (const provider of providers) provider.destroy();
    await left.dispose();
    await right?.dispose();
  }
});

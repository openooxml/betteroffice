import { expect, spyOn, test } from 'bun:test';
import { PipedCollaboration, type PipeHostTransport } from './main';
import type { PipeErrorCode, PipeOp, PipeOutput, PipePeer } from './protocol';
import { clockOrigin } from './protocol';
import { createCollaborationPipeHost } from './worker';

type Event = Parameters<Parameters<PipeHostTransport['onEvent']>[0]>[0];
type CodedError = Error & { code: PipeErrorCode };

function rebasePeer(peer: PipePeer, offset: number): PipePeer {
  return { ...peer, lastSeenAt: peer.lastSeenAt + offset,
    cursorMovedAt: peer.cursorMovedAt === 0 ? 0 : peer.cursorMovedAt + offset };
}

function setup(options: { maxPendingBytes?: number; send?: (data: Uint8Array) => boolean } = {}) {
  const sent: Uint8Array[] = [];
  const posts: PipeOp[][] = [];
  const errors: CodedError[] = [];
  let listener: ((event: Event) => void) | null = null;
  let receive!: (out: readonly PipeOutput[]) => void;
  let lost!: (reason: string) => void;
  let disconnects = 0;
  const transport: PipeHostTransport = {
    connect() {},
    disconnect() { disconnects++; },
    send(data) { sent.push(data); return options.send ? options.send(data) : true; },
    onEvent(next) { listener = next; return () => { listener = null; }; },
  };
  const makeError = (code: PipeErrorCode, message: string) => Object.assign(new Error(message), { code });
  const piped = new PipedCollaboration<CodedError>({
    open(onOut, onLost) {
      receive = onOut;
      lost = onLost;
      return { post: (ops) => { posts.push(ops); return new Promise(() => {}); }, close() {} };
    },
  }, transport, { worker: {}, maxPendingBytes: options.maxPendingBytes ?? 1 << 20, hasUser: true, rebasePeer, makeError,
    normalize: (code, context) => makeError(code, context) });
  piped.onError((error) => errors.push(error));
  piped.connect();
  const emit = (event: Event) => listener!(event);
  return { piped, sent, posts, errors, emit, receive: (out: PipeOutput[]) => receive(out), lose: (reason: string) => lost(reason),
    disconnects: () => disconnects };
}

test('pipe peers rebase worker timestamps without extending the flag window on heartbeat or delivery', () => {
  const workerOffset = 20_000;
  const workerTimeOrigin = performance.timeOrigin + workerOffset;
  const clock = spyOn(performance, 'now').mockReturnValue(workerOffset + 150);
  const pipe = setup();
  const peer: PipePeer = { clientId: 8, clock: 1, user: { name: 'B', color: '#000' }, cursor: null,
    lastSeenAt: 100, cursorMovedAt: 80, inferredCursor: null };
  const received: Array<readonly PipePeer[]> = [];
  pipe.piped.onPeers(peers => received.push(peers));
  const deliver = (peers: PipePeer[]) => {
    const output = { do: 'peers' as const, peers, timeOrigin: workerTimeOrigin, through: 3 };
    pipe.receive([output]);
  };
  try {
    deliver([peer]);
    expect(pipe.piped.peers[0]).toEqual({ ...peer, lastSeenAt: workerOffset + 100, cursorMovedAt: workerOffset + 80 });
    expect(received.at(-1)).toBe(pipe.piped.peers);
    expect(performance.now() - pipe.piped.peers[0]!.lastSeenAt).toBe(50);
    expect(performance.now() - pipe.piped.peers[0]!.cursorMovedAt).toBe(70);
    expect(peer).toMatchObject({ lastSeenAt: 100, cursorMovedAt: 80 });

    clock.mockReturnValue(workerOffset + 3_200);
    deliver([{ ...peer, clock: 2, lastSeenAt: 3_000 }]);
    expect(pipe.piped.peers[0]).toMatchObject({ lastSeenAt: workerOffset + 3_000, cursorMovedAt: workerOffset + 80 });
    expect(performance.now() - pipe.piped.peers[0]!.cursorMovedAt).toBeGreaterThanOrEqual(3_000);

    deliver([{ ...peer, clock: 3, lastSeenAt: 3_100, cursorMovedAt: 3_100 }]);
    expect(pipe.piped.peers[0]!.cursorMovedAt).toBe(workerOffset + 3_100);
    expect(performance.now() - pipe.piped.peers[0]!.cursorMovedAt).toBe(100);

    deliver([{ ...peer, cursorMovedAt: 0 }]);
    expect(pipe.piped.peers[0]!.cursorMovedAt).toBe(0);
  } finally { pipe.piped.destroy(); clock.mockRestore(); }
});

test('worker peer output carries its time origin without changing provider timestamps', () => {
  let publish!: (peers: readonly PipePeer[]) => void;
  const host = createCollaborationPipeHost({
    clientId: 1,
    read: { encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0) },
    onUpdate: () => () => {},
    merge: () => ({ applied: null }),
    createProvider: () => ({
      connect() {}, disconnect() {}, destroy() {}, setCursor() {},
      onStatus: () => () => {}, onError: () => () => {},
      onPeers(listener: typeof publish) { publish = listener; listener([]); return () => {}; },
    }),
    arrive: () => ({ settle() {}, fail() {} }),
    leave: () => null,
    transportError: error => error,
  });
  const peer: PipePeer = { clientId: 8, clock: 1, user: { name: 'B', color: '#000' }, cursor: null,
    lastSeenAt: 100, cursorMovedAt: 80, inferredCursor: null };
  try {
    host.run({ provider: 1, ops: [{ op: 'create', seq: 1, options: {} }] });
    publish([peer]);
    expect(host.drain()?.batches).toEqual([{ provider: 1,
      out: [{ do: 'peers', peers: [peer], timeOrigin: performance.timeOrigin, through: 1 }] }]);
    expect(peer).toMatchObject({ lastSeenAt: 100, cursorMovedAt: 80 });
  } finally { host.dispose(); }
});

test('worker sends from before a transport close are dropped', () => {
  const pipe = setup();
  pipe.emit({ type: 'open' });
  pipe.emit({ type: 'close' });
  pipe.receive([{ do: 'send', data: new Uint8Array([1]).buffer, through: 3 }]);
  expect(pipe.sent).toHaveLength(0);
  pipe.receive([{ do: 'send', data: new Uint8Array([2]).buffer, through: 4 }]);
  expect(pipe.sent.map((frame) => frame[0])).toEqual([2]);
});

test('received frames past maxPendingBytes while the worker is busy fail the connection', () => {
  const pipe = setup({ maxPendingBytes: 250 });
  pipe.emit({ type: 'open' });
  pipe.emit({ type: 'message', data: new Uint8Array(100) });
  pipe.emit({ type: 'message', data: new Uint8Array(0) });
  expect(pipe.errors).toHaveLength(0);
  pipe.emit({ type: 'message', data: new Uint8Array(0) });
  expect(pipe.errors.map((error) => error.code)).toEqual(['backpressure']);
  expect(pipe.piped.status).toBe('disconnected');
});

test('a lone received frame may use the whole maxPendingBytes', () => {
  const pipe = setup({ maxPendingBytes: 100 });
  pipe.emit({ type: 'open' });
  pipe.emit({ type: 'message', data: new Uint8Array(100) });
  expect(pipe.errors).toHaveLength(0);
  pipe.emit({ type: 'message', data: new Uint8Array(0) });
  expect(pipe.errors.map((error) => error.code)).toEqual(['backpressure']);
});

test('a lone received frame past maxPendingBytes fails the connection', () => {
  const pipe = setup({ maxPendingBytes: 100 });
  pipe.emit({ type: 'open' });
  pipe.emit({ type: 'message', data: new Uint8Array(101) });
  expect(pipe.errors.map((error) => error.code)).toEqual(['backpressure']);
});

test('a drain raised inside send does not resend the queued frame', () => {
  let refuse = true;
  let emit!: (event: Event) => void;
  const pipe = setup({ send: () => {
    if (refuse) { refuse = false; return false; }
    emit({ type: 'drain' });
    return true;
  } });
  emit = pipe.emit;
  pipe.emit({ type: 'open' });
  pipe.receive([{ do: 'send', data: new Uint8Array([7]).buffer, through: 3 }]);
  expect(pipe.piped.pendingBytes).toBe(1);
  pipe.emit({ type: 'drain' });
  expect(pipe.sent).toHaveLength(2);
  expect(pipe.piped.pendingBytes).toBe(0);
});

test('a retired replica sends the cached leave before handing the transport back', () => {
  const pipe = setup();
  pipe.emit({ type: 'open' });
  pipe.receive([
    { do: 'status', status: 'connected', synced: true, through: 3 },
    { do: 'leave', data: new Uint8Array([9]).buffer, through: 3 },
  ]);
  pipe.lose('retired');
  expect(pipe.sent.map((frame) => frame[0])).toEqual([9]);
  expect(pipe.disconnects()).toBe(1);
});

test('a transport close reports disconnected, unsynced and no peers at once, and the worker echo adds nothing', () => {
  const pipe = setup();
  const statuses: string[] = [];
  const peerCounts: number[] = [];
  pipe.piped.onStatus(({ status, synced }) => statuses.push(`${status}:${synced}`));
  pipe.piped.onPeers((peers) => peerCounts.push(peers.length));
  pipe.emit({ type: 'open' });
  const peer = { clientId: 8, clock: 1, user: { name: 'B', color: '#000' }, cursor: null, lastSeenAt: 1, cursorMovedAt: 1,
    inferredCursor: null };
  pipe.receive([
    { do: 'status', status: 'connected', synced: true, through: 3 },
    { do: 'peers', peers: [peer], timeOrigin: performance.timeOrigin, through: 3 },
  ]);
  pipe.emit({ type: 'close' });
  expect([pipe.piped.status, pipe.piped.synced, pipe.piped.peers.length]).toEqual(['disconnected', false, 0]);
  pipe.receive([
    { do: 'status', status: 'connected', synced: true, through: 3 },
    { do: 'status', status: 'disconnected', synced: false, through: 4 },
    { do: 'peers', peers: [], timeOrigin: performance.timeOrigin, through: 4 },
  ]);
  expect(statuses).toEqual(['connected:true', 'disconnected:false']);
  expect(peerCounts).toEqual([0, 1, 0]);
});

test('a reconnect from a status listener reaches the worker after the close that prompted it', async () => {
  const kinds: string[] = [];
  let listener: ((event: Event) => void) | null = null;
  let openOnConnect = false;
  const transport: PipeHostTransport = {
    connect() { if (openOnConnect) listener!({ type: 'open' }); },
    disconnect() {},
    send() { return true; },
    onEvent(next) { listener = next; return () => { listener = null; }; },
  };
  const makeError = (code: PipeErrorCode, message: string) => Object.assign(new Error(message), { code });
  const piped = new PipedCollaboration<CodedError>({
    open() {
      return { post: (ops) => {
        kinds.push(...ops.map((op) => op.op === 'event' ? op.event.type : op.op));
        return Promise.resolve();
      }, close() {} };
    },
  }, transport, { worker: {}, maxPendingBytes: 1 << 20, hasUser: true, rebasePeer, makeError, normalize: (code, context) => makeError(code, context) });
  piped.connect();
  listener!({ type: 'open' });
  await new Promise((resolve) => setTimeout(resolve, 0));
  piped.onStatus(({ status }) => {
    if (status !== 'disconnected' || openOnConnect) return;
    openOnConnect = true;
    piped.connect();
  });
  kinds.length = 0;
  listener!({ type: 'close' });
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(kinds).toEqual(['close', 'connect', 'open']);
  expect(piped.status).toBe('connecting');
});

test('the clock origin follows the clock the provider reads when timeOrigin or performance is missing', () => {
  const now = spyOn(performance, 'now').mockReturnValue(1_500);
  const wall = spyOn(Date, 'now').mockReturnValue(1_700_000_000_000);
  const original = globalThis.performance;
  const own = Object.getOwnPropertyDescriptor(performance, 'timeOrigin');
  try {
    expect(clockOrigin()).toBe(performance.timeOrigin);
    Object.defineProperty(performance, 'timeOrigin', { value: undefined, configurable: true });
    expect(clockOrigin()).toBe(1_700_000_000_000 - 1_500);
    Object.defineProperty(globalThis, 'performance', { value: undefined, configurable: true, writable: true });
    expect(clockOrigin()).toBe(0);
  } finally {
    Object.defineProperty(globalThis, 'performance', { value: original, configurable: true, writable: true });
    if (own) Object.defineProperty(performance, 'timeOrigin', own);
    else delete (performance as { timeOrigin?: number }).timeOrigin;
    now.mockRestore();
    wall.mockRestore();
  }
});

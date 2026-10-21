import { expect, test } from 'bun:test';
import { PipedCollaboration, type PipeHostTransport } from './main';
import type { PipeErrorCode, PipeOp, PipeOutput } from './protocol';

type Event = Parameters<Parameters<PipeHostTransport['onEvent']>[0]>[0];
type CodedError = Error & { code: PipeErrorCode };

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
  }, transport, { worker: {}, maxPendingBytes: options.maxPendingBytes ?? 1 << 20, hasUser: true, makeError,
    normalize: (code, context) => makeError(code, context) });
  piped.onError((error) => errors.push(error));
  piped.connect();
  const emit = (event: Event) => listener!(event);
  return { piped, sent, posts, errors, emit, receive: (out: PipeOutput[]) => receive(out), lose: (reason: string) => lost(reason),
    disconnects: () => disconnects };
}

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
    { do: 'peers', peers: [peer], through: 3 },
  ]);
  pipe.emit({ type: 'close' });
  expect([pipe.piped.status, pipe.piped.synced, pipe.piped.peers.length]).toEqual(['disconnected', false, 0]);
  pipe.receive([
    { do: 'status', status: 'connected', synced: true, through: 3 },
    { do: 'status', status: 'disconnected', synced: false, through: 4 },
    { do: 'peers', peers: [], through: 4 },
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
  }, transport, { worker: {}, maxPendingBytes: 1 << 20, hasUser: true, makeError, normalize: (code, context) => makeError(code, context) });
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

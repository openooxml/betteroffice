import { afterEach, describe, expect, mock, setSystemTime, spyOn, test } from 'bun:test';
import * as Y from 'yjs';
import { decodeMessages, encodeSyncStep1, encodeSyncStep2, encodeUpdate } from '../../../packages/docx/src/collaboration/protocol';
import { documentFrame, rehydrate } from './fixtures';
import { updateKey } from '../src/persistence';

mock.module('cloudflare:workers', () => ({
  DurableObject: class {
    constructor(protected readonly ctx: unknown, protected readonly env: unknown) {}
  },
}));
const { CollaborationRoom } = await import('../src/index');

function createSocket() {
  const socket = { readyState: 1, send: mock((_: unknown) => {}), close: mock((..._: unknown[]) => { socket.readyState = 3; }) };
  return socket;
}
type Socket = ReturnType<typeof createSocket>;
(globalThis as { WebSocketPair?: unknown }).WebSocketPair = function () {
  return { 0: createSocket(), 1: createSocket() };
};
const UPGRADE = new Request('https://relay.test/room/a', { headers: { Upgrade: 'websocket' } });

function createRoom(seed: Iterable<[string, unknown]> = []) {
  const sender = createSocket();
  const peer = createSocket();
  const sockets: Socket[] = [sender, peer];
  const rows = new Map<string, unknown>(seed);
  const pending: Promise<unknown>[] = [];
  const alarms: number[] = [];
  let failWrite = false;
  let failDelete = false;
  let beforeCommit: (() => Promise<void>) | undefined;
  let initialization = Promise.resolve();
  const storageFor = (data: Map<string, unknown>) => ({
    get: async (key: string | string[]) => Array.isArray(key) ? new Map(key.filter(k => data.has(k)).map(k => [k, data.get(k)])) : data.get(key),
    list: async ({ prefix }: { prefix: string }) => new Map([...data].filter(([key]) => key.startsWith(prefix)).sort()),
    put: async (key: string | Record<string, unknown>, value?: unknown) => {
      if (failWrite) throw new Error('Injected storage failure');
      const entries = typeof key === 'string' ? [[key, value] as const] : Object.entries(key);
      expect(entries.length).toBeLessThanOrEqual(128);
      for (const [name, bytes] of entries) {
        if (bytes instanceof Uint8Array) expect(bytes.length).toBeLessThanOrEqual(64 * 1024);
        data.set(name, structuredClone(bytes));
      }
    },
    delete: async (key: string | string[]) => {
      if (failDelete) throw new Error('Injected delete failure');
      const keys = Array.isArray(key) ? key : [key];
      expect(keys.length).toBeLessThanOrEqual(128);
      for (const name of keys) data.delete(name);
      return keys.length;
    },
  });
  const storage = {
    ...storageFor(rows),
    transaction: async (run: (tx: ReturnType<typeof storageFor>) => Promise<void>) => {
      const staged = new Map(rows);
      await run(storageFor(staged));
      await beforeCommit?.();
      rows.clear();
      for (const [key, value] of staged) rows.set(key, value);
    },
    getAlarm: async () => null,
    setAlarm: async (time: number) => { alarms.push(time); },
    deleteAll: async () => { rows.clear(); },
  };
  const state = {
    storage,
    blockConcurrencyWhile: (initialize: () => Promise<void>) => { initialization = initialize(); },
    acceptWebSocket: (socket: Socket) => sockets.push(socket),
    getWebSockets: () => sockets,
    waitUntil: (promise: Promise<unknown>) => pending.push(promise),
  };
  const room = new CollaborationRoom(state as never, {} as never);
  return { room, sender, peer, sockets, rows, pending, initialization, storage, alarms,
    fail: (operation: 'put' | 'delete' = 'put') => { if (operation === 'put') failWrite = true; else failDelete = true; },
    onCommit: (callback?: () => Promise<void>) => { beforeCommit = callback; },
  };
}
type Harness = ReturnType<typeof createRoom>;
async function flush(harness: Harness) { await Promise.all(harness.pending); }
function send(harness: Harness, frame: Uint8Array) { harness.room.webSocketMessage(harness.sender as never, frame.buffer.slice(frame.byteOffset, frame.byteOffset + frame.byteLength) as ArrayBuffer); }
async function join(harness: Harness) {
  await harness.room.fetch(UPGRADE);
  return harness.sockets.at(-1)!;
}
function frames(socket: Socket): Uint8Array[] {
  return socket.send.mock.calls.map(([bytes]) => bytes).filter((bytes): bytes is Uint8Array => bytes instanceof Uint8Array);
}

afterEach(() => { setSystemTime(); mock.restore(); });

describe('CollaborationRoom', () => {
  test('persists document updates before broadcasting and skips awareness', async () => {
    const h = createRoom();
    await h.initialization;
    const document = documentFrame();
    let release!: () => void;
    const gate = new Promise<void>(resolve => { release = resolve; });
    h.onCommit(() => gate);
    send(h, document);
    await Promise.resolve();
    expect(h.peer.send).not.toHaveBeenCalled();
    release();
    await flush(h);
    expect(h.rows.get(updateKey(0))).toEqual(document);
    expect(frames(h.peer)).toEqual([document]);
    expect(h.sender.send).not.toHaveBeenCalled();
    send(h, Uint8Array.of(1, 1, 12));
    await flush(h);
    expect(h.rows.size).toBe(1);
    expect(frames(h.peer)).toHaveLength(2);
  });

  test('does not treat a disconnected peer as a persistence failure', async () => {
    const h = createRoom();
    await h.initialization;
    h.peer.readyState = 3;
    h.peer.send.mockImplementation(() => { throw new Error('Socket already closed'); });
    send(h, documentFrame());
    await flush(h);
    h.room.webSocketClose(h.peer as never, 1000, '', true);
    await h.room.checkpoint();
    expect(h.rows.has('checkpoint')).toBe(true);
    expect(h.sender.close).not.toHaveBeenCalled();
  });

  test('refuses malformed payloads and auth without persisting or broadcasting them', async () => {
    const h = createRoom();
    await h.initialization;
    for (const frame of [Uint8Array.of(0x80), Uint8Array.of(0, 2, 1, 255), Uint8Array.of(2, 0, 1, 65)]) send(h, frame);
    await flush(h);
    expect(h.sender.close).toHaveBeenCalledTimes(3);
    expect(h.peer.send).not.toHaveBeenCalled();
    expect(h.rows.size).toBe(0);
  });

  test('checkpoints after 512 updates and rehydrates after every client leaves', async () => {
    const h = createRoom();
    await h.initialization;
    const doc = new Y.Doc();
    doc.on('update', update => send(h, encodeUpdate(update)));
    for (let i = 0; i < 520; i++) doc.getText('body').insert(i, String(i % 10));
    await flush(h);
    expect([...h.rows.keys()].filter(key => key.startsWith('update:'))).toHaveLength(8);
    expect(h.rows.has('checkpoint')).toBe(true);
    const restarted = createRoom(h.rows);
    await restarted.initialization;
    restarted.sockets.length = 0;
    const joined = await join(restarted);
    const restored = rehydrate(frames(joined));
    expect(restored.getText('body').toString()).toBe(doc.getText('body').toString());
    expect(decodeMessages(frames(joined).at(-1)!)[0].type).toBe('sync-step-1');
    doc.destroy(); restored.destroy();
  });

  test('requests the initial seed and synchronizes an offline client against the checkpoint', async () => {
    const h = createRoom();
    await h.initialization;
    const source = rehydrate([documentFrame('seed')]);
    const offline = rehydrate([documentFrame('seed')]);
    const joined = await join(h);
    const query = decodeMessages(frames(joined)[0])[0];
    expect(query.type).toBe('sync-step-1');
    if (query.type !== 'sync-step-1') throw new Error('Missing initial handshake');
    send(h, encodeSyncStep2(Y.encodeStateAsUpdate(source, query.stateVector)));
    await flush(h);
    await h.room.checkpoint();
    source.getText('body').insert(4, ' online');
    send(h, encodeUpdate(Y.encodeStateAsUpdate(source)));
    await flush(h);
    await h.room.checkpoint();
    offline.getText('body').insert(0, 'offline ');
    const restart = createRoom(h.rows);
    await restart.initialization;
    send(restart, encodeSyncStep1(Y.encodeStateVector(offline)));
    await flush(restart);
    for (const frame of frames(restart.sender)) for (const message of decodeMessages(frame)) {
      if (message.type === 'sync-step-2') Y.applyUpdate(offline, message.update);
    }
    send(restart, encodeUpdate(Y.encodeStateAsUpdate(offline)));
    await flush(restart);
    await restart.room.checkpoint();
    const final = createRoom(restart.rows);
    await final.initialization;
    const restored = rehydrate(frames(await join(final)));
    expect(restored.getText('body').toString()).toBe('offline seed online');
    for (const doc of [source, offline, restored]) doc.destroy();
  });

  test('stores large checkpoints in bounded chunks and restores all bytes', async () => {
    const h = createRoom();
    await h.initialization;
    const text = 'large document '.repeat(15000);
    send(h, documentFrame(text));
    await flush(h);
    expect([...h.rows.keys()].filter(key => key.startsWith('checkpoint:')).length).toBeGreaterThan(2);
    const restart = createRoom(h.rows);
    await restart.initialization;
    const doc = rehydrate(frames(await join(restart)));
    expect(doc.getText('body').toString()).toBe(text);
    doc.destroy();
  });

  test('serializes updates arriving during checkpoint persistence', async () => {
    const h = createRoom();
    await h.initialization;
    send(h, documentFrame('first', 1));
    await flush(h);
    let release!: () => void;
    let started!: () => void;
    const ready = new Promise<void>(resolve => { started = resolve; });
    const gate = new Promise<void>(resolve => { release = resolve; });
    h.onCommit(async () => { started(); await gate; });
    const checkpoint = h.room.checkpoint();
    await ready;
    send(h, documentFrame('second', 2));
    expect(h.rows.has('checkpoint')).toBe(false);
    release();
    await checkpoint;
    await flush(h);
    const restart = createRoom(h.rows);
    await restart.initialization;
    const doc = rehydrate(frames(await join(restart)));
    expect(doc.getText('body').toString()).toBe('firstsecond');
    doc.destroy();
  });

  test('keeps committed data after a failed checkpoint and refuses uncommitted broadcasts', async () => {
    const h = createRoom();
    await h.initialization;
    send(h, documentFrame());
    await flush(h);
    const before = structuredClone(h.rows);
    h.fail('delete');
    const log = spyOn(console, 'error').mockImplementation(() => {});
    await expect(h.room.checkpoint()).rejects.toThrow('Room storage unavailable');
    expect(h.rows).toEqual(before);
    expect(h.sender.close).toHaveBeenCalledWith(1011, 'Room storage unavailable; reconnect');
    send(h, documentFrame('unsaved', 2));
    await flush(h);
    expect(frames(h.peer)).toHaveLength(1);
    expect((await h.room.fetch(UPGRADE)).status).toBe(503);
    const restart = createRoom(h.rows);
    await restart.initialization;
    expect(rehydrate(frames(await join(restart))).getText('body').toString()).toBe('hello');
    expect(log).toHaveBeenCalled();
  });

  test('migrates the retained and legacy logs without dropping their updates', async () => {
    const h = createRoom([[updateKey(7), documentFrame('new', 2)], ['updates', [documentFrame('old', 1)]]]);
    await h.initialization;
    expect(h.rows.has('updates')).toBe(false);
    expect(h.rows.has(updateKey(7))).toBe(false);
    const restart = createRoom(h.rows);
    await restart.initialization;
    expect(rehydrate(frames(await join(restart))).getText('body').toString()).toBe('oldnew');
    send(restart, documentFrame('tail', 3));
    await flush(restart);
    expect(restart.rows.has(updateKey(9))).toBe(true);
  });

  test('refuses incomplete checkpoint storage instead of serving a partial document', async () => {
    const h = createRoom();
    await h.initialization;
    send(h, documentFrame());
    await flush(h);
    await h.room.checkpoint();
    h.rows.delete('checkpoint:0000');
    await expect(createRoom(h.rows).initialization).rejects.toThrow('Incomplete relay checkpoint');
  });

  test('retains the existing idle-room expiration policy', async () => {
    const start = Date.UTC(2026, 0, 1);
    setSystemTime(start);
    const h = createRoom();
    await h.initialization;
    send(h, documentFrame());
    await flush(h);
    await h.room.checkpoint();
    expect(h.alarms).toEqual([start + 86400000]);
    setSystemTime(start + 86400000);
    await h.room.alarm();
    expect(h.rows.has('checkpoint')).toBe(true);
    h.sockets.length = 0;
    await h.room.alarm();
    expect(h.rows.size).toBe(0);
    expect(frames(await join(h))).toEqual([encodeSyncStep1(Uint8Array.of(0))]);
  });
});

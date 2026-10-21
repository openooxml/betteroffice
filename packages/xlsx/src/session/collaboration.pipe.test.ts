import { beforeAll, expect, test } from 'bun:test';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { collaborationPipeFor } from '../collaboration/asyncReplica';
import { CollaborationProvider } from '../collaboration/provider';
import { decodeAwarenessUpdate, decodeMessages, encodeAwarenessUpdate, encodeSyncStep2, encodeUpdate } from '../collaboration/protocol';
import type { CollaborationError, CollaborationProviderOptions, CollaborationTransport, CollaborationTransportEvent } from '../collaboration/types';
import { openWorkbook } from '../wasm/loader';
import type { WorkbookAck, WorkbookSharedCommand } from './commandProtocol';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';

let bytes: Uint8Array;
beforeAll(async () => { bytes = (await loadWorkbookSessionFixtures()).fixture; });

class Wire implements CollaborationTransport {
  readonly frames: Uint8Array[] = [];
  readonly calls: string[] = [];
  readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();
  accept = true;
  connect() { this.calls.push('connect'); this.emit({ type: 'open' }); }
  disconnect() { this.calls.push('disconnect'); }
  send(data: Uint8Array) { this.frames.push(data.slice()); return this.accept; }
  onEvent(listener: (event: CollaborationTransportEvent) => void) {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }
  emit(event: CollaborationTransportEvent) { for (const listener of [...this.listeners]) listener(event); }
}

async function until(done: () => boolean | Promise<boolean>) {
  for (let turn = 0; turn < 1000; turn++) {
    if (await done()) return;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error('Timed out waiting for collaboration');
}

async function open(options: CollaborationProviderOptions = {}) {
  const requests: WorkbookSharedCommand[] = [];
  const session = await createTestWorkbookSession(bytes, (transport) => ({ ...transport, listen(listener) {
    return transport.listen((message) => {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'edit') {
        requests.push(structuredClone(message.args[0] as WorkbookSharedCommand));
      }
      listener(message);
    });
  } }), { collaborative: true, clientId: 8101,
    initialViewport: { width: 800, height: 600, zoom: 1, dpr: 1, requestGeneration: 0 } });
  const wire = new Wire();
  const provider = new CollaborationProvider(session.replica!, wire, options);
  const errors: CollaborationError[] = [];
  provider.onError((error) => errors.push(error));
  provider.connect();
  await until(() => provider.status === 'connected' && wire.frames.length > 0);
  return { session, requests, wire, provider, errors };
}

async function sync(room: Awaited<ReturnType<typeof open>>, peer: ReturnType<typeof openWorkbook>) {
  const count = room.requests.length;
  room.wire.emit({ type: 'message', data: encodeSyncStep2(peer.encodeStateAsUpdate()) });
  await until(() => room.provider.synced && room.session.throughSeq > count);
}

test('a worker provider merges a transport burst in one collabPipe command, version and frame', async () => {
  const room = await open();
  const peer = openWorkbook(bytes, { collaborative: true, clientId: 8102 });
  try {
    expect(collaborationPipeFor(room.session.replica!)).toBeDefined();
    await sync(room, peer);
    const updates: Uint8Array[] = [];
    peer.onUpdate((update, origin) => { if (origin === 'local') updates.push(update); });
    for (let col = 1; col <= 24; col++) peer.editCell(0, 2, col, `burst ${col}`);
    expect(updates).toHaveLength(24);
    const before = { requests: room.requests.length, version: room.session.state.version };
    const acks: WorkbookAck[] = [];
    room.session.onAck((ack) => acks.push(ack));
    for (const update of updates) room.wire.emit({ type: 'message', data: encodeUpdate(update) });
    await until(() => acks.length > 0);
    expect(room.requests.slice(before.requests).map((request) => request.kind)).toEqual(['collabPipe']);
    const command = room.requests[before.requests].payload.payload;
    expect('ops' in command && command.ops.filter((op) => op.op === 'event')).toHaveLength(24);
    expect(room.session.state.version).toBe(before.version + 1);
    expect(acks).toHaveLength(1);
    expect(acks[0].statuses[0].result).toMatchObject({ remoteUpdate: true });
    expect(acks.filter((ack) => ack.frame)).toHaveLength(1);
    expect(acks[0].frame?.version).toBe(acks[0].version);
    const cells = (await room.session.call.cellInputs(0, 'B3:Y3')).cells;
    expect(cells[0].map((cell) => cell.input)).toEqual(Array.from({ length: 24 }, (_, index) => `burst ${index + 1}`));
    expect(room.session.editState.canUndo).toBe(false);
    expect(room.errors).toEqual([]);
  } finally { room.provider.destroy(); peer.dispose(); await room.session.dispose(); }
});

test('worker status, sync and awareness reach the facade, including a timed cursor broadcast and leave', async () => {
  const room = await open({ user: { name: 'Ada' } });
  const peer = openWorkbook(bytes, { collaborative: true, clientId: 8102 });
  const cursor = { sheet: 'Sheet1', anchor: { row: 2, col: 1 }, head: { row: 4, col: 3 } };
  const seen: number[] = [];
  room.provider.onAwareness((peers) => seen.push(peers.length));
  try {
    expect(room.provider.synced).toBe(false);
    await sync(room, peer);
    room.wire.emit({ type: 'message', data: encodeAwarenessUpdate([{ clientId: peer.clientId, clock: 1,
      state: { user: { name: 'Bo', color: '#0B57D0' }, cursor } }]) });
    await until(() => room.provider.peers.length === 1);
    expect(room.provider.peers[0]).toMatchObject({ clientId: peer.clientId, user: { name: 'Bo' }, cursor });
    expect(seen).toEqual([0, 1]);
    room.provider.setCursor(cursor);
    const awareness = () => room.wire.frames.flatMap((frame) => decodeMessages(frame))
      .flatMap((message) => message.type === 'awareness' ? decodeAwarenessUpdate(message.update) : []);
    await until(() => awareness().some((update) => update.state?.cursor?.head.row === 4));
    room.provider.destroy();
    expect(room.provider.status).toBe('destroyed');
    const states = awareness();
    expect(states[states.length - 1]?.state).toBeNull();
    expect(room.wire.listeners.size).toBe(0);
    expect(room.errors).toEqual([]);
  } finally { room.provider.destroy(); peer.dispose(); await room.session.dispose(); }
});

test('a refused update fails the provider with replica error and leaves the worker usable', async () => {
  const room = await open();
  const peer = openWorkbook(bytes, { collaborative: true, clientId: 8102 });
  try {
    await sync(room, peer);
    const version = await room.session.call.version();
    room.wire.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(0xff)) });
    expect(room.errors).toEqual([]);
    await until(() => room.errors.length > 0);
    expect(room.errors[0]).toMatchObject({ code: 'replica' });
    expect(room.errors[0].message).toContain('Failed to apply remote update');
    expect(room.provider.status).toBe('disconnected');
    expect(room.provider.synced).toBe(false);
    expect(room.wire.listeners.size).toBe(0);
    expect(await room.session.call.version()).toBe(version);
    expect(room.session.failure).toBeUndefined();
    expect((await room.session.editSession!.submit('editCell', { payload: { args: [0, 2, 1, 'still alive'] } }).ack).status).toBe('applied');
    expect((await room.session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('still alive');
  } finally { room.provider.destroy(); peer.dispose(); await room.session.dispose(); }
});

test('disconnect and reconnect resync through the same worker provider', async () => {
  const room = await open();
  const peer = openWorkbook(bytes, { collaborative: true, clientId: 8102 });
  try {
    await sync(room, peer);
    room.provider.disconnect();
    expect(room.provider.status).toBe('disconnected');
    expect(room.provider.synced).toBe(false);
    expect(room.wire.listeners.size).toBe(0);
    peer.editCell(0, 2, 1, 'offline');
    room.provider.connect();
    expect(room.provider.status).toBe('connecting');
    await until(() => room.provider.status === 'connected');
    await sync(room, peer);
    expect((await room.session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('offline');
    expect(room.wire.calls).toEqual(['connect', 'disconnect', 'connect']);
    room.provider.destroy();
    room.provider.connect();
    expect(room.provider.status).toBe('destroyed');
    expect(room.wire.calls).toEqual(['connect', 'disconnect', 'connect', 'disconnect']);
    expect(room.errors).toEqual([]);
  } finally { room.provider.destroy(); peer.dispose(); await room.session.dispose(); }
});

test('transport backpressure stays on main and drain retries the queued frames', async () => {
  const room = await open();
  try {
    room.wire.accept = false;
    room.provider.setCursor({ sheet: 'Sheet1', anchor: { row: 1, col: 1 }, head: { row: 1, col: 1 } });
    await until(() => room.provider.pendingBytes > 0);
    const frames = room.wire.frames.length;
    room.wire.accept = true;
    room.wire.emit({ type: 'drain' });
    expect(room.provider.pendingBytes).toBe(0);
    expect(room.wire.frames.length).toBeGreaterThan(frames);
    expect(room.errors).toEqual([]);
  } finally { room.provider.destroy(); await room.session.dispose(); }
});

test('maxPendingBytes fails the connection with backpressure', async () => {
  const room = await open({ maxPendingBytes: 0 });
  try {
    room.wire.accept = false;
    room.provider.setCursor({ sheet: 'Sheet1', anchor: { row: 1, col: 1 }, head: { row: 1, col: 1 } });
    await until(() => room.errors.length > 0);
    expect(room.errors[0]).toMatchObject({ code: 'backpressure' });
    expect(room.provider.status).toBe('disconnected');
    expect(room.provider.pendingBytes).toBe(0);
    expect(room.wire.listeners.size).toBe(0);
    expect(room.session.failure).toBeUndefined();
  } finally { room.provider.destroy(); await room.session.dispose(); }
});

test('worker failure retires the pipe and a new worker resyncs', async () => {
  let fail!: () => void;
  const session = await createTestWorkbookSession(bytes, (transport) => ({ ...transport, post(message, transfer) {
    transport.post(message, transfer);
    if (isHostMessage(message) && message.kind === 'reply') fail = () => transport.post({ protocol: 1,
      kind: 'failure', code: 'crash', message: 'Worker lost' });
  } }), { collaborative: true, clientId: 8201 });
  const wire = new Wire();
  const provider = new CollaborationProvider(session.replica!, wire);
  const errors: CollaborationError[] = [];
  provider.onError((error) => errors.push(error));
  let replacement: Awaited<ReturnType<typeof open>> | undefined;
  const peer = openWorkbook(bytes, { collaborative: true, clientId: 8102 });
  try {
    provider.connect();
    await until(() => provider.status === 'connected');
    fail();
    await until(() => errors.length > 0);
    expect(errors[0]).toMatchObject({ code: 'replica' });
    expect(provider.status).toBe('disconnected');
    expect(provider.synced).toBe(false);
    expect(wire.listeners.size).toBe(0);
    peer.editCell(0, 2, 1, 'resynced');
    replacement = await open();
    await sync(replacement, peer);
    expect((await replacement.session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('resynced');
    expect(replacement.errors).toEqual([]);
  } finally {
    provider.destroy(); peer.dispose(); await session.dispose();
    replacement?.provider.destroy(); await replacement?.session.dispose();
  }
});

test('a refused update after applied ones in a burst keeps them and fails only the connection', async () => {
  const room = await open();
  const peer = openWorkbook(bytes, { collaborative: true, clientId: 8102 });
  try {
    await sync(room, peer);
    const updates: Uint8Array[] = [];
    peer.onUpdate((update, origin) => { if (origin === 'local') updates.push(update); });
    peer.editCell(0, 2, 1, 'partial');
    room.wire.emit({ type: 'message', data: encodeUpdate(updates[0]) });
    room.wire.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(0xff)) });
    await until(() => room.errors.length > 0);
    expect(room.errors[0]).toMatchObject({ code: 'replica' });
    expect(room.wire.listeners.size).toBe(0);
    expect(room.session.failure).toBeUndefined();
    expect((await room.session.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('partial');
  } finally { room.provider.destroy(); peer.dispose(); await room.session.dispose(); }
});

import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/yrs/wasm/index';
import { preloadOpcWasm } from '@betteroffice/docx/wasm/opc';
import { residentWorkerFactory } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { ResidentSourceWorkerClient, createDocxWorkerEditSession, createResidentCommandConnection } from '@betteroffice/docx/peerless';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { CollaborationProvider, type CollaborationError, type CollaborationReplica, type CollaborationTransport,
  type CollaborationTransportEvent } from '@betteroffice/docx/collaboration';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';

const TIMEOUT = 30_000;
const generated = resolve(import.meta.dir, '../../../../../docx/src/wasm/generated');
let start: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(generated, 'edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(generated, 'opc/ooxml_opc_bg.wasm'))));
  start = await residentWorkerFactory({ transferBuffers: true });
});

const documentBytes = () => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
  ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')],
])));

async function open(clientId: number, initialUpdate?: Uint8Array) {
  const worker = start(clientId);
  const client = new ResidentSourceWorkerClient(worker);
  await client.openSource({ bytes: documentBytes(), collaboration: { clientId, initialUpdate } });
  const engine = worker.sessions[0]!;
  const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 4 };
  const initial = await client.commandInit({ session: `replica-${clientId}`, incarnation: 1, epoch: 1,
    readOnly: false, selection: { anchor: at, focus: at },
    fonts: [new Uint8Array(readFileSync(resolve(import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')))],
    layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} }),
    layoutWithRegions: true, layoutRevision: 1, extras: '{}',
  });
  const session = createDocxWorkerEditSession({ source: documentBytes(),
    connection: await createResidentCommandConnection(client, initial),
    echo: { paint() {}, drop() {}, present() {} }, maxRecoveryAttempts: 0,
    recovery: { reopen: () => Promise.reject(new Error('Unexpected reopen')) },
  });
  const replica = createPeerlessCollaborationReplica(client, { session }, clientId);
  const merges = { single: 0, chunks: 0 };
  const applyUpdate = engine.applyUpdate, applyUpdateChunks = engine.applyUpdateChunks;
  Object.assign(engine, {
    applyUpdate(update: Uint8Array) { merges.single++; return applyUpdate.call(engine, update); },
    applyUpdateChunks(updates: readonly Uint8Array[]) { merges.chunks++; return applyUpdateChunks.call(engine, updates); },
  });
  const edits = () => worker.requests.filter(type => type === 'edit').length;
  const close = async () => { replica.retire(); await session.dispose(); };
  return { worker, engine, session, replica, merges, edits, close };
}

class Wire implements CollaborationTransport {
  peer?: Wire;
  readonly sent: Uint8Array[] = [];
  readonly calls: string[] = [];
  accept = true;
  readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();
  connect() { this.calls.push('connect'); }
  disconnect() { this.calls.push('disconnect'); }
  send(data: Uint8Array) {
    this.sent.push(data);
    if (!this.accept) return false;
    queueMicrotask(() => this.peer?.emit({ type: 'message', data }));
    return true;
  }
  onEvent(listener: (event: CollaborationTransportEvent) => void) {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }
  emit(event: CollaborationTransportEvent) { for (const listener of [...this.listeners]) listener(event); }
}

async function until(done: () => boolean | Promise<boolean>, what = 'condition') {
  for (let tick = 0; tick < 4000; tick++) {
    if (await done()) return;
    await Bun.sleep(5);
  }
  throw new Error(`Timed out waiting for ${what}`);
}

const text = (session: Pick<YrsSession, 'paragraphs'>) => session.paragraphs('body').map(paragraph => paragraph.text).join('\n');

function pair(left: CollaborationReplica, right: CollaborationReplica, options: { left?: object; right?: object } = {}) {
  const a = new Wire(), b = new Wire();
  a.peer = b; b.peer = a;
  const providers = [new CollaborationProvider(left, a, options.left), new CollaborationProvider(right, b, options.right)] as const;
  const errors: CollaborationError[] = [];
  for (const provider of providers) { provider.onError(error => errors.push(error)); provider.connect(); }
  a.emit({ type: 'open' }); b.emit({ type: 'open' });
  return { wires: [a, b] as const, providers, errors, synced: () => providers.every(provider => provider.synced) };
}

test('a burst of remote edits merges in one command and one engine merge', async () => {
  const a = await open(9411);
  const peer = await createYrsSession({ clientId: 9412 });
  peer.loadState(a.engine.encodeState());
  const room = pair(a.replica, peer);
  try {
    await until(room.synced, 'sync');
    const before = { edits: a.edits(), ...a.merges };
    const paraId = peer.paragraphs('body')[0]!.paraId;
    for (let index = 0; index < 40; index++) peer.insertText({ story: 'body', paraId, offset: 0 }, String.fromCharCode(97 + (index % 26)));
    await until(() => text(a.engine) === text(peer), 'the burst to merge');
    expect(room.errors).toEqual([]);
    expect(a.edits() - before.edits).toBe(1);
    expect(a.merges.single + a.merges.chunks - before.single - before.chunks).toBe(1);
  } finally { for (const provider of room.providers) provider.destroy(); peer.destroy(); await a.close(); }
}, TIMEOUT);

test('the main thread forwards frames undecoded; the worker fails the connection on a malformed one', async () => {
  const a = await open(9413);
  const wire = new Wire();
  const provider = new CollaborationProvider(a.replica, wire);
  const errors: CollaborationError[] = [];
  provider.onError(error => errors.push(error));
  try {
    provider.connect();
    wire.emit({ type: 'open' });
    await until(() => provider.status === 'connected' && wire.sent.length > 0, 'the handshake');
    wire.emit({ type: 'message', data: Uint8Array.of(0xff, 0xff, 0xff) });
    expect(errors).toEqual([]);
    expect(provider.status).toBe('connected');
    await until(() => errors.length > 0, 'the worker to refuse the frame');
    expect(errors[0]).toMatchObject({ code: 'protocol' });
    expect(errors[0]!.message).toContain('Invalid collaboration frame');
    expect(provider.status).toBe('disconnected');
    expect(wire.calls).toEqual(['connect', 'disconnect']);
    expect(wire.listeners.size).toBe(0);
  } finally { provider.destroy(); await a.close(); }
}, TIMEOUT);

test('local edits leave from the worker, remote edits merge, and undo reverts only the local edit', async () => {
  const a = await open(9414);
  const peer = await createYrsSession({ clientId: 9415 });
  peer.loadState(a.engine.encodeState());
  const room = pair(a.replica, peer);
  try {
    await until(room.synced, 'sync');
    expect((await a.session.submit('insertText', { text: ' local' }).ack).status).toBe('applied');
    await until(() => text(peer) === 'Seed local', 'the local edit to reach the peer');
    peer.insertText({ story: 'body', paraId: peer.paragraphs('body')[0]!.paraId, offset: 0 }, 'remote ');
    await until(() => text(a.engine) === 'remote Seed local', 'the remote edit to merge');
    expect((await a.session.submit('undo', {}).ack).status).toBe('applied');
    await until(() => text(peer) === 'remote Seed' && text(a.engine) === 'remote Seed', 'the undo to reach the peer');
    expect(room.errors).toEqual([]);
  } finally { for (const provider of room.providers) provider.destroy(); peer.destroy(); await a.close(); }
}, TIMEOUT);

test('a reconnect resyncs what both sides edited while apart', async () => {
  const a = await open(9416);
  const peer = await createYrsSession({ clientId: 9417 });
  peer.loadState(a.engine.encodeState());
  const room = pair(a.replica, peer);
  const [mine, theirs] = room.providers;
  try {
    await until(room.synced, 'sync');
    mine.disconnect();
    expect(mine.status).toBe('disconnected');
    expect(room.wires[0].listeners.size).toBe(0);
    peer.insertText({ story: 'body', paraId: peer.paragraphs('body')[0]!.paraId, offset: 0 }, 'offline ');
    expect((await a.session.submit('insertText', { text: ' apart' }).ack).status).toBe('applied');
    expect(text(a.engine)).toBe('Seed apart');
    mine.connect();
    room.wires[0].emit({ type: 'open' });
    room.wires[1].emit({ type: 'close' });
    room.wires[1].emit({ type: 'open' });
    await until(() => room.synced() && text(a.engine) === text(peer), 'the resync');
    expect(text(peer)).toBe('offline Seed apart');
    expect(theirs.synced).toBe(true);
    expect(room.errors).toEqual([]);
  } finally { for (const provider of room.providers) provider.destroy(); peer.destroy(); await a.close(); }
}, TIMEOUT);

test('presence runs in the worker: peers, cursors and the leave on destroy reach the other side', async () => {
  const a = await open(9418);
  const b = await open(9419, a.engine.encodeStateAsUpdate());
  const room = pair(a.replica, b.replica, { left: { user: { name: 'Ada' } }, right: { user: { name: 'Bo' } } });
  const [left, right] = room.providers;
  try {
    await until(() => room.synced() && left.peers.length === 1 && right.peers.length === 1, 'presence');
    expect(right.peers[0]).toMatchObject({ clientId: 9418, user: { name: 'Ada' } });
    const paraId = a.engine.paragraphs('body')[0]!.paraId;
    const sticky = a.engine.encodeStickyPosition({ story: 'body', paraId, offset: 2 });
    left.setCursor({ story: 'body', anchor: sticky.encoded, head: sticky.encoded });
    await until(() => right.peers[0]?.cursor !== null && right.peers[0]?.cursor !== undefined, 'the cursor');
    expect([...right.peers[0]!.cursor!.head]).toEqual([...sticky.encoded]);
    left.destroy();
    expect(left.status).toBe('destroyed');
    expect(room.wires[0].calls).toEqual(['connect', 'disconnect']);
    await until(() => right.peers.length === 0, 'the leave');
    expect(room.errors).toEqual([]);
  } finally { for (const provider of room.providers) provider.destroy(); await a.close(); await b.close(); }
}, TIMEOUT);

test('refused frames wait on the main thread for drain', async () => {
  const a = await open(9420);
  const wire = new Wire();
  wire.accept = false;
  const provider = new CollaborationProvider(a.replica, wire);
  try {
    provider.connect();
    wire.emit({ type: 'open' });
    await until(() => provider.pendingBytes > 0, 'a refused frame');
    const refused = wire.sent.length;
    wire.accept = true;
    wire.emit({ type: 'drain' });
    expect(provider.pendingBytes).toBe(0);
    expect(wire.sent.length).toBeGreaterThan(refused);
  } finally { provider.destroy(); await a.close(); }
}, TIMEOUT);

test('a retired replica hands the transport back at once and reports only to a provider still alive', async () => {
  const a = await open(9421);
  const wire = new Wire();
  const provider = new CollaborationProvider(a.replica, wire);
  const errors: CollaborationError[] = [];
  provider.onError(error => errors.push(error));
  try {
    provider.connect();
    wire.emit({ type: 'open' });
    await until(() => provider.status === 'connected', 'the connection');
    a.replica.retire();
    expect(wire.listeners.size).toBe(0);
    expect(wire.calls).toEqual(['connect', 'disconnect']);
    expect(errors).toEqual([]);
    await until(() => errors.length > 0, 'the retirement report');
    expect(errors[0]).toMatchObject({ code: 'replica' });
    expect(provider.status).toBe('disconnected');
  } finally { provider.destroy(); await a.close(); }
}, TIMEOUT);

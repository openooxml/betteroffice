import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/yrs/wasm/index';
import { preloadOpcWasm } from '@betteroffice/docx/wasm/opc';
import { residentWorkerFactory } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { ResidentSourceWorkerClient, createDocxWorkerEditSession, createResidentCommandConnection } from '@betteroffice/docx/peerless';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { CollaborationProvider, type CollaborationError, type CollaborationReplica, type CollaborationTransport,
  type CollaborationTransportEvent } from '@betteroffice/docx/collaboration';
import { decodeMessages } from '@betteroffice/docx/collaboration/protocol';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';

const TIMEOUT = 20_000;
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
  return { engine, session, replica };
}

test('local typing submitted while a refused merge is pending still applies', async () => {
  const a = await open(9301);
  try {
    const merged = a.replica.async.applyUpdate(new Uint8Array([255]));
    const typing = a.session.submit('insertText', { text: ' typed' });
    await expect(merged).rejects.toThrow('refused');
    expect(await typing.ack).toMatchObject({ status: 'applied' });
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed typed');
  } finally { a.replica.retire(); await a.session.dispose(); }
}, TIMEOUT);

test('handshake reads contain local input admitted before them', async () => {
  const a = await open(9302);
  const peer = await createYrsSession({ clientId: 9399 });
  try {
    peer.loadState(a.engine.encodeState());
    const vector = peer.encodeStateVector();
    const typing = a.session.submit('insertText', { text: ' typed' });
    const diff = await a.replica.async.encodeStateAsUpdate(vector);
    peer.applyUpdate(diff);
    expect(peer.paragraphs('body')[0]!.text).toBe('Seed typed');
    expect((await typing.ack).status).toBe('applied');
    const next = a.session.submit('insertText', { text: '!' });
    expect(await a.replica.async.encodeStateVector()).toEqual(a.engine.encodeStateVector());
    expect((await next.ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed typed!');
  } finally { peer.destroy(); a.replica.retire(); await a.session.dispose(); }
}, TIMEOUT);

const WORDS = ['one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten'];

const singleStruct = (update: Uint8Array) => update[0] === 1 && update[1] === 1;

class Loopback implements CollaborationTransport {
  peer?: Loopback;
  readonly sent: Uint8Array[] = [];
  private readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();
  connect() {}
  disconnect() {}
  send(data: Uint8Array) {
    this.sent.push(data);
    queueMicrotask(() => this.peer?.emit({ type: 'message', data }));
    return true;
  }
  onEvent(listener: (event: CollaborationTransportEvent) => void) {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }
  emit(event: CollaborationTransportEvent) { for (const listener of [...this.listeners]) listener(event); }
}

/** Syncs a point-to-point pair whose seeder chunks its answer; `updates` counts the Updates it sent before its SyncStep2. */
async function sync(seeder: CollaborationReplica, joiner: CollaborationReplica) {
  const seederWire = new Loopback(), joinerWire = new Loopback();
  seederWire.peer = joinerWire;
  joinerWire.peer = seederWire;
  const providers = [new CollaborationProvider(seeder, seederWire, { syncChunkBytes: 32 }), new CollaborationProvider(joiner, joinerWire)];
  const errors: CollaborationError[] = [];
  for (const provider of providers) { provider.onError(error => errors.push(error)); provider.connect(); }
  seederWire.emit({ type: 'open' });
  joinerWire.emit({ type: 'open' });
  for (let tick = 0; tick < 2000 && !errors.length && !providers.every(provider => provider.synced); tick++) await Bun.sleep(5);
  const synced = providers.every(provider => provider.synced);
  for (const provider of providers) provider.destroy();
  const messages = seederWire.sent.flatMap(frame => decodeMessages(frame));
  const answer = messages.slice(0, messages.findIndex(message => message.type === 'sync-step-2'));
  return { errors, synced, updates: answer.filter(message => message.type === 'update').length };
}

test('a worker replica answers a chunked diff that merges like the whole one', async () => {
  const a = await open(9303);
  const whole = await createYrsSession({ clientId: 9398 });
  const chunked = await createYrsSession({ clientId: 9397 });
  try {
    for (const word of WORDS) expect((await a.session.submit('insertText', { text: ` ${word}` }).ack).status).toBe('applied');
    const vector = whole.encodeStateVector();
    const chunks = await a.replica.async.encodeStateAsUpdateChunks!(vector, 64);
    expect(chunks.length).toBeGreaterThan(1);
    for (const chunk of chunks) expect(chunk.byteLength <= 64 || singleStruct(chunk)).toBe(true);
    whole.applyUpdate(await a.replica.async.encodeStateAsUpdate(vector));
    chunked.applyUpdateChunks(chunks);
    const texts = (session: Pick<typeof whole, 'paragraphs'>) => session.paragraphs('body').map(paragraph => paragraph.text);
    expect(texts(whole)).toEqual(texts(a.engine));
    expect(texts(chunked)).toEqual(texts(whole));
    expect(chunked.encodeStateVector()).toEqual(whole.encodeStateVector());
    expect(chunked.encodeStateVector()).toEqual(a.engine.encodeStateVector());
  } finally { whole.destroy(); chunked.destroy(); a.replica.retire(); await a.session.dispose(); }
}, TIMEOUT);

test('a worker joiner merges a chunked answer from a yrs seeder', async () => {
  const a = await open(9304);
  const seeder = await createYrsSession({ clientId: 9396 });
  try {
    seeder.loadState(a.engine.encodeState());
    const paraId = seeder.paragraphs('body')[0]!.paraId;
    for (const word of WORDS) seeder.insertText({ story: 'body', paraId, offset: 0 }, `${word} `);
    const room = await sync(seeder, a.replica);
    expect(room.errors).toEqual([]);
    expect(room.synced).toBe(true);
    expect(room.updates).toBeGreaterThan(1);
    expect(a.engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(seeder.paragraphs('body').map(paragraph => paragraph.text));
    expect(a.engine.encodeStateVector()).toEqual(seeder.encodeStateVector());
  } finally { seeder.destroy(); a.replica.retire(); await a.session.dispose(); }
}, TIMEOUT);

test('a yrs joiner merges a chunked answer from a worker seeder', async () => {
  const a = await open(9305);
  const joiner = await createYrsSession({ clientId: 9395 });
  try {
    await a.session.submit('insertText', { text: ' typed' }).ack;
    const room = await sync(a.replica, joiner);
    expect(room.errors).toEqual([]);
    expect(room.synced).toBe(true);
    expect(room.updates).toBeGreaterThan(1);
    expect(joiner.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed typed']);
    expect(joiner.encodeStateVector()).toEqual(a.engine.encodeStateVector());
  } finally { joiner.destroy(); a.replica.retire(); await a.session.dispose(); }
}, TIMEOUT);

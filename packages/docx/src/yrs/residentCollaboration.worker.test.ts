import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadOpcWasm } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentSourceWorkerClient } from './residentSourceWorkerClient';
import { createResidentCommandConnection } from './residentCommandConnection';
import { createDocxWorkerEditSession } from './createDocxWorkerEditSession';
import { createYrsSession } from './index';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';
import type { Document } from '../types/document';

let start: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await openCommandEngine();
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'))));
  start = await residentWorkerFactory();
});
afterEach(closeCommandEngines);

function documentBytes() {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')],
  ])));
}

async function parsedDocument(): Promise<Document> {
  const parser = await createYrsSession({ clientId: 9299 });
  try {
    parser.openDocx(documentBytes(), true);
    return parser.materializeDocx()!;
  } finally { parser.destroy(); }
}

async function openCommands(clientId: number, initialUpdate?: Uint8Array, document?: Document) {
  const worker = start(clientId);
  const client = new ResidentSourceWorkerClient(worker);
  await client.openSource({ ...(document ? { document } : { bytes: documentBytes() }), collaboration: { clientId, initialUpdate } });
  const engine = worker.sessions[0]!;
  const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 4 };
  const initial = await client.commandInit({ session: `replica-${clientId}`, incarnation: 1, epoch: 1,
    readOnly: false, selection: { anchor: at, focus: at },
    fonts: [new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))],
    layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} }),
    layoutWithRegions: true, layoutRevision: 1, extras: '{}',
  });
  return { worker, client, engine, initial };
}

async function open(clientId: number, initialUpdate?: Uint8Array, document?: Document) {
  const { worker, client, engine, initial } = await openCommands(clientId, initialUpdate, document);
  const session = createDocxWorkerEditSession({ source: documentBytes(),
    connection: await createResidentCommandConnection(client, initial),
    echo: { paint() {}, drop() {}, present() {} }, maxRecoveryAttempts: 0,
    recovery: { reopen: () => Promise.reject(new Error('Unexpected reopen')) },
  });
  return { worker, client, engine, initial, session };
}

const remoteUpdate = (target: { engine: { encodeStateVector(): Uint8Array } }, source: { encodeStateAsUpdate(vector?: Uint8Array): Uint8Array }) =>
  source.encodeStateAsUpdate(target.engine.encodeStateVector()).slice().buffer;
const merge = { lane: 'collab', fence: true } as const;

for (const input of ['bytes', 'document'] as const) {
  test(`a ${input} source joins a room from the shared state without seeding its own stories`, async () => {
    const a = await open(9201);
    const b = await open(9202, a.engine.encodeState(), input === 'document' ? await parsedDocument() : undefined);
    const pushed: Uint8Array[] = [];
    const off = b.client.onCollabUpdates(event => pushed.push(...event.updates));
    try {
      expect(b.engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed', 'Second']);
      expect(b.engine.paragraphs('body')).toEqual(a.engine.paragraphs('body'));
      expect((await b.session.submit('insertText', { text: 'B' }).ack).status).toBe('applied');
      expect(pushed.length).toBeGreaterThan(0);
      for (const update of pushed) await a.session.submit('applyRemoteUpdate', { update: update.slice().buffer }, { lane: 'collab', fence: true }).ack;
      expect(a.engine.paragraphs('body')).toEqual(b.engine.paragraphs('body'));
    } finally { off(); a.session.dispose(); b.session.dispose(); }
  }, TIMEOUT);
}

test('two worker command sessions converge with reordered and duplicated updates', async () => {
  const a = await open(9201);
  const b = await open(9202, a.engine.encodeState());
  const left: ArrayBuffer[] = [], right: ArrayBuffer[] = [];
  const offA = a.client.onCollabUpdates(event => left.push(...event.updates.map(u => u.slice().buffer)));
  const offB = b.client.onCollabUpdates(event => right.push(...event.updates.map(u => u.slice().buffer)));
  try {
    await Promise.all([a.session.submit('insertText', { text: 'A' }).ack, b.session.submit('insertText', { text: 'B' }).ack]);
    await a.session.submit('insertText', { text: 'a' }).ack;
    for (const update of [...left].reverse()) await b.session.submit('applyRemoteUpdate', { update }, { lane: 'collab', fence: true }).ack;
    for (const update of [...right, ...right]) await a.session.submit('applyRemoteUpdate', { update }, { lane: 'collab', fence: true }).ack;
    for (const update of left) await b.session.submit('applyRemoteUpdate', { update }, { lane: 'collab', fence: true }).ack;
    expect(a.engine.paragraphs('body')).toEqual(b.engine.paragraphs('body'));
    expect(a.engine.paragraphs('body')[0]!.text).toContain('A');
    expect(a.engine.paragraphs('body')[0]!.text).toContain('B');
  } finally { offA(); offB(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

test('update push precedes command ack and queued typing survives a merge', async () => {
  const a = await open(9201);
  const b = await open(9202, a.engine.encodeState());
  const events: string[] = [];
  a.client.onCollabUpdates(() => { throw new Error('Listener failure'); });
  const off = a.client.onCollabUpdates(() => { events.push('update'); });
  try {
    await a.session.submit('insertText', { text: 'A' }).ack.then(() => events.push('ack'));
    expect(events).toEqual(['update', 'ack']);
    await b.session.submit('insertText', { text: 'B' }).ack;
    const merge = a.session.submit('applyRemoteUpdate', { update: b.engine.encodeStateAsUpdate(a.engine.encodeStateVector()).slice().buffer }, { lane: 'collab', fence: true });
    const typing = a.session.submit('insertText', { text: '!' }, { after: merge });
    expect((await merge.ack).status).toBe('applied');
    expect((await typing.ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toContain('!');
    const pending = a.session.submit('insertText', { text: 'discard' });
    a.session.dispose();
    expect((await pending.ack).status).toBe('cancelled');
  } finally { off(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

test('presence reads are tied to the covering frame and unresolved anchors are marked', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    const paragraphs = a.engine.paragraphs('body');
    const at = { story: 'body', paraId: paragraphs[1]!.paraId, offset: 3 };
    await a.session.submit('select', { selection: { anchor: at, focus: at } }).ack;
    const selected = await a.client.requestFullFrame();
    const cursor = a.engine.encodeSelection()!;
    const version = a.engine.proposalEngine.version();
    const read = { kind: 'collabPresence' as const, peers: [{ clientId: 9202, clock: 1, cursor }], pages: [1, 2] as [number, number],
      expectVersion: version, expectFrameEpoch: selected.frame!.revision, expectLayoutRevision: selected.frame!.payload.layoutRevision };
    expect((await a.client.documentRead(read)).value?.peers[0]).toMatchObject({ status: 'ok', caret: { pageIndex: 1 } });
    expect((await a.client.documentRead({ ...read, expectFrameEpoch: 999 })).value).toBeNull();
    expect((await a.client.documentRead({ ...read, peers: [{ clientId: 9202, clock: 2,
      cursor: { ...cursor, anchor: new Uint8Array([255]) } }] })).value?.peers[0]?.status).toBe('unresolved');
    remote.loadState(a.engine.encodeState());
    const table = remote.insertTable({ story: 'body', paraId: paragraphs[0]!.paraId, offset: 0 }, 1, 1);
    const story = table.createdStoryIds[0]!;
    const paraId = remote.paragraphs(story)[0]!.paraId;
    remote.insertText({ story, paraId, offset: 0 }, 'Anchor');
    await a.session.submit('applyRemoteUpdate', { update: remote.encodeStateAsUpdate(a.engine.encodeStateVector()).slice().buffer }, { lane: 'collab', fence: true }).ack;
    remote.setSelection({ story, paraId, offset: 3 });
    const deletedCursor = remote.encodeSelection()!;
    remote.deleteTable(table.table);
    await a.session.submit('applyRemoteUpdate', { update: remote.encodeStateAsUpdate(a.engine.encodeStateVector()).slice().buffer }, { lane: 'collab', fence: true }).ack;
    const frame = await a.client.requestFullFrame();
    const current = a.engine.proposalEngine.version();
    expect((await a.client.documentRead({ ...read, peers: [{ clientId: 9203, clock: 3, cursor: deletedCursor }], pages: [0, 2], expectVersion: current,
      expectFrameEpoch: frame.frame!.revision, expectLayoutRevision: frame.frame!.payload.layoutRevision })).value?.peers[0]?.status).toBe('unresolved');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('inferred remote carets rebase through a later local insert', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    const paraId = a.engine.paragraphs('body')[0]!.paraId;
    remote.insertText({ story: 'body', paraId, offset: 4 }, 'A');
    await a.session.submit('applyRemoteUpdate', { update: remote.encodeStateAsUpdate(a.engine.encodeStateVector()).slice().buffer }, { lane: 'collab', fence: true }).ack;
    const presence = async () => {
      const frame = (await a.client.requestFullFrame()).frame!;
      return (await a.client.documentRead({ kind: 'collabPresence', peers: [{ clientId: 9203, clock: 1,
        inferred: { story: 'body', paraId, offset: 5 } }], pages: [0, 1], expectVersion: a.engine.proposalEngine.version(),
        expectFrameEpoch: frame.revision, expectLayoutRevision: frame.payload.layoutRevision })).value?.peers[0]?.caret;
    };
    const before = await presence();
    const at = { story: 'body', paraId, offset: 0 };
    await a.session.submit('insertText', { text: 'prefix', selection: { anchor: at, focus: at } }).ack;
    const after = await presence();
    expect(before).not.toBeNull();
    expect(after!.x).toBeGreaterThan(before!.x);
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('a read-only replica merges remote edits and still refuses local ones', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    expect((await a.session.submit('setReadOnly', { readOnly: true }).ack).status).toBe('noop');
    expect(a.session.state.editState?.readOnly).toBe(true);
    const version = a.session.state.version;
    remote.loadState(a.engine.encodeState());
    remote.insertText({ story: 'body', paraId: a.engine.paragraphs('body')[0]!.paraId, offset: 4 }, ' remote');
    expect((await a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge).ack).status).toBe('applied');
    expect(a.session.state.version).toBe(version + 1);
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed remote');
    expect(await a.session.submit('insertText', { text: 'local' }).ack).toMatchObject({ status: 'refused', reason: 'read-only' });
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed remote');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('a merge relayouts on the command path and its ack frame holds the merged text', async () => {
  const a = await openCommands(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    const paraId = a.engine.paragraphs('body')[0]!.paraId;
    remote.insertText({ story: 'body', paraId, offset: 4 }, ' merged');
    const ack = await a.client.edit({ session: a.initial.session, incarnation: 1, epoch: 1, seq: 1,
      basis: { epoch: 1, version: 0 }, kind: 'applyRemoteUpdate', payload: { update: remoteUpdate(a, remote) } });
    expect(ack).toMatchObject({ version: 1, throughSeq: 1, outcomes: [{ seq: 1, status: 'applied' }] });
    const version = a.engine.proposalEngine.version();
    expect(ack.frame!.payload.documentVersion).toBe(version);
    const cursor = (offset: number) => {
      remote.setSelection({ story: 'body', paraId, offset });
      return remote.encodeSelection()!;
    };
    const read = await a.client.documentRead({ kind: 'collabPresence', pages: [0, 1], expectVersion: version,
      expectFrameEpoch: ack.frame!.revision, expectLayoutRevision: ack.frame!.payload.layoutRevision,
      peers: [{ clientId: 9203, clock: 1, cursor: cursor(4) }, { clientId: 9204, clock: 1, cursor: cursor(11) }] });
    const [before, after] = read.value!.peers;
    expect(before).toMatchObject({ status: 'ok', caret: { pageIndex: 0 } });
    expect(after).toMatchObject({ status: 'ok', caret: { pageIndex: 0 } });
    expect(after!.caret!.x).toBeGreaterThan(before!.caret!.x);
  } finally { remote.destroy(); a.client.destroy(); }
}, TIMEOUT);

test('host proposal mutations are refused in a collaborating command session and push nothing', async () => {
  const a = await open(9201);
  const pushed: Uint8Array[] = [];
  const off = a.client.onCollabUpdates(event => pushed.push(...event.updates));
  try {
    const before = a.engine.encodeStateVector();
    await expect(a.client.proposal({ kind: 'propose', request: { expectVersion: a.engine.proposalEngine.version(), proposals: [{
      id: 'proposal', paragraph: { kind: 'persisted', story: { kind: 'body', partUri: '/word/document.xml' }, paraId: a.engine.paragraphs('body')[0]!.paraId },
      suggest: { author: 'Host', date: '2026-10-13T00:00:00Z' }, op: 'insertText', at: 'end', text: ' proposed',
    }] } })).rejects.toThrow(/has not laid out its document|Legacy mutations are unavailable/);
    expect(pushed).toEqual([]);
    expect(a.engine.encodeStateVector()).toEqual(before);
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed');
  } finally { off(); a.session.dispose(); }
}, TIMEOUT);

test('a collaborating session refuses the unordered remote update request', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    remote.insertText({ story: 'body', paraId: a.engine.paragraphs('body')[0]!.paraId, offset: 4 }, ' remote');
    await expect(a.client.applyRemoteUpdate(new Uint8Array(remoteUpdate(a, remote)))).rejects.toThrow('applyRemoteUpdate command');
    expect(a.client.hasFailed()).toBe(false);
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed');
    expect((await a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge).ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed remote');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

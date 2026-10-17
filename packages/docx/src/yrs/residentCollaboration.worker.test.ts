import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentSourceWorkerClient } from './residentSourceWorkerClient';
import { createResidentCommandConnection } from './residentCommandConnection';
import { createDocxWorkerEditSession } from './createDocxWorkerEditSession';
import { createYrsSession } from './index';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';
import { MAX_SYNC_CHUNKS_BYTES, MAX_UPDATE_BYTES } from './residentCommandActions';
import { encodeUpdate, MAX_COLLABORATION_FRAME_BYTES, stateVectorClock } from '../collaboration/protocol';
import { applyFrameDeltaOwned, decodeFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import type { Document } from '../types/document';
import type { DocxProposalInput } from './proposals';
import * as wasm from './wasm/index';
import { CollaborationProvider } from '../collaboration/provider';
import { createAsyncCollaborationReplica } from '../collaboration/asyncReplica';
import type { CollaborationError, CollaborationTransport, CollaborationTransportEvent } from '../collaboration/types';

let start: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await openCommandEngine();
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'))));
  start = await residentWorkerFactory({ transferBuffers: true });
});
afterEach(closeCommandEngines);

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const SEED_BODY = '<w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>';

function documentBytes(header?: string, body = SEED_BODY, sectPr = '<w:sectPr/>') {
  const parts: PartsMap = new Map();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    (header === undefined ? '' : '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>') + '</Types>');
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
  if (header !== undefined) {
    set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/></Relationships>`);
    set('word/header1.xml', `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${header}</w:t></w:r></w:p></w:hdr>`);
  }
  set('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"${header === undefined ? '' : ` xmlns:r="${REL}"`}><w:body>${body}` +
    (header === undefined ? sectPr : '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr>') + '</w:body></w:document>');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function parsedDocument(): Promise<Document> {
  const parser = await createYrsSession({ clientId: 9299 });
  try {
    parser.openDocx(documentBytes(), true);
    return parser.materializeDocx()!;
  } finally { parser.destroy(); }
}

async function openCommands(clientId: number, initialUpdate?: Uint8Array, document?: Document, bytes = documentBytes(), worker = start(clientId)) {
  const client = new ResidentSourceWorkerClient(worker);
  const opened = await client.openSource({ ...(document ? { document } : { bytes }), collaboration: { clientId, initialUpdate } });
  const engine = worker.sessions[0]!;
  const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 4 };
  const initial = await client.commandInit({ session: `replica-${clientId}`, incarnation: 1, epoch: 1,
    readOnly: false, selection: { anchor: at, focus: at },
    fonts: [new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))],
    layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} }),
    layoutWithRegions: true, layoutRevision: 1, extras: '{}',
  });
  return { worker, client, engine, initial, opened };
}

async function open(clientId: number, initialUpdate?: Uint8Array, document?: Document, bytes = documentBytes(), worker = start(clientId)) {
  const { client, engine, initial, opened } = await openCommands(clientId, initialUpdate, document, bytes, worker);
  const session = createDocxWorkerEditSession({ source: bytes,
    connection: await createResidentCommandConnection(client, initial),
    echo: { paint() {}, drop() {}, present() {} }, maxRecoveryAttempts: 0,
    recovery: { reopen: () => Promise.reject(new Error('Unexpected reopen')) },
  });
  return { worker, client, engine, initial, session, opened };
}

const remoteUpdate = (target: { engine: { encodeStateVector(): Uint8Array } }, source: { encodeStateAsUpdate(vector?: Uint8Array): Uint8Array }) =>
  source.encodeStateAsUpdate(target.engine.encodeStateVector()).slice().buffer;
const merge = { lane: 'collab', fence: true } as const;
const clocks = (vector: Uint8Array) => {
  let at = 0;
  const read = () => { let value = 0; for (let scale = 1; ; scale *= 128) { const byte = vector[at++]!; value += (byte & 127) * scale; if (byte < 128) return value; } };
  return Object.fromEntries(Array.from({ length: read() }, () => [read(), read()]));
};

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

test('reserved protocol client ids are refused', async () => {
  for (const clientId of [Number.MAX_SAFE_INTEGER - 1, Number.MAX_SAFE_INTEGER]) await expect(createYrsSession({ clientId })).rejects.toThrow(RangeError);
  const session = await createYrsSession({ clientId: Number.MAX_SAFE_INTEGER - 2 });
  expect(session.clientId).toBe(Number.MAX_SAFE_INTEGER - 2);
  session.destroy();
}, TIMEOUT);

test('a pinned id opens under a fresh id when the opened state lags the clock it emitted', async () => {
  const bytes = documentBytes();
  const a = await open(9501, undefined, undefined, bytes, start());
  const before = a.engine.encodeState();
  const clocks: number[] = [];
  const off = a.client.onCollabUpdates(event => { if (event.clock !== undefined) clocks.push(event.clock); });
  try {
    expect(a.engine.clientId).toBe(9501);
    expect((await a.session.submit('insertText', { text: ' typed while pinned' }).ack).status).toBe('applied');
    const clock = clocks.at(-1)!;
    expect(clock).toBeGreaterThan(0);
    expect(clock).toBe(stateVectorClock(a.engine.encodeStateVector(), 9501));
    for (const [initialUpdate, kept] of [[before, false], [a.engine.encodeState(), true]] as const) {
      const worker = start();
      const client = new ResidentSourceWorkerClient(worker);
      try {
        const opened = await client.openSource({ bytes, collaboration: { clientId: 9501, initialUpdate, clientClock: clock } });
        expect(opened.clientId === 9501).toBe(kept);
        expect(worker.sessions.map(session => session.clientId)).toEqual(kept ? [9501] : [9501, opened.clientId!]);
      } finally { client.destroy(); }
    }
  } finally { off(); a.session.dispose(); }
}, TIMEOUT);

test('a pinned id that seeds again opens under a fresh id and converges with a peer of its first seed', async () => {
  const seed = async (body: string, clientClock?: number) => {
    const worker = start();
    const client = new ResidentSourceWorkerClient(worker);
    const opened = await client.openSource({ bytes: documentBytes(undefined, body), collaboration: { clientId: 9511, clientClock } });
    return { client, opened, engine: worker.sessions.at(-1)! };
  };
  const first = await seed(SEED_BODY);
  const peer = await createYrsSession({ clientId: 9512 });
  const second = await seed(`${SEED_BODY}<w:p><w:r><w:t>A longer replacement seed</w:t></w:r></w:p>`, stateVectorClock(first.opened.stateVector, 9511));
  try {
    peer.loadState(first.engine.encodeState());
    expect(second.opened.clientId).not.toBe(9511);
    expect(second.engine.clientId).toBe(second.opened.clientId!);
    peer.applyUpdate(second.engine.encodeStateAsUpdate(peer.encodeStateVector()));
    second.engine.applyUpdate(peer.encodeStateAsUpdate(second.engine.encodeStateVector()));
    expect(peer.paragraphs('body')).toEqual(second.engine.paragraphs('body'));
  } finally { first.client.destroy(); second.client.destroy(); peer.destroy(); }
}, TIMEOUT);

test('a pinned session reports its own clocks held at open and merged from a peer', async () => {
  const bytes = documentBytes();
  const a = await open(9521, undefined, undefined, bytes, start());
  const checkpoint = a.engine.encodeState();
  const held = stateVectorClock(a.engine.encodeStateVector(), 9521);
  expect((await a.session.submit('insertText', { text: ' later' }).ack).status).toBe('applied');
  const b = await open(9521, checkpoint, undefined, bytes, start());
  const clocks: number[] = [];
  const off = b.client.onCollabUpdates(event => { if (event.clock !== undefined) clocks.push(event.clock); });
  try {
    expect(held).toBeGreaterThan(0);
    expect((await b.session.submit('applyRemoteUpdate', { update: remoteUpdate(b, a.engine) }, merge).ack).status).toBe('applied');
    expect(clocks).toEqual([stateVectorClock(a.engine.encodeStateVector(), 9521)]);
    expect(clocks[0]).toBeGreaterThan(held);
    expect(b.opened.clientClock).toBe(held);
  } finally { off(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

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

const ROOM_BODY = ['ALPHA BETA', 'GAMMA DELTA', 'EPSILON ZETA', 'ETA THETA']
  .map(text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('');

test('a worker replica survives a remote update that arrives before the update it depends on and converges once it does', async () => {
  const failures: string[] = [];
  const awaited: string[] = [];
  for (const steps of [['split', 'undo'], ['deleteAcross', 'split', 'undo'], ['section', 'split', 'undo'],
    ['section', 'deleteAcross', 'split', 'undo'], ['table', 'deleteAcross', 'split', 'undo']]) {
    const bytes = documentBytes(undefined, ROOM_BODY);
    const a = await open(9201, undefined, undefined, bytes);
    const b = await open(9202, a.engine.encodeState(), undefined, bytes);
    let sent: ArrayBuffer[] = [];
    const back: ArrayBuffer[] = [];
    const offA = a.client.onCollabUpdates(event => sent.push(...event.updates.map(u => u.slice().buffer)));
    const offB = b.client.onCollabUpdates(event => back.push(...event.updates.map(u => u.slice().buffer)));
    const point = (index: number, offset: number) => ({ story: 'body', paraId: a.engine.paragraphs('body')[index]!.paraId, offset });
    const caret = (index: number, offset: number) => ({ selection: { anchor: point(index, offset), focus: point(index, offset) } });
    const label = steps.join('+');
    try {
      const groups: ArrayBuffer[][] = [];
      for (const step of steps) {
        sent = [];
        const [kind, payload] = step === 'split' ? ['paragraphBreak', caret(0, 3)]
          : step === 'deleteAcross' ? ['deleteBackward', { selection: { anchor: point(0, 3), focus: point(1, Math.min(4, a.engine.paragraphs('body')[1]!.text.length)) } }]
            : step === 'section' ? ['insertSectionBreak', { ...caret(1, 2), type: 'continuous' }]
              : step === 'table' ? ['insertTable', { ...caret(1, 0), rows: 2, columns: 2 }] : ['undo', {}];
        const status = (await a.session.submit(kind as 'undo', payload as Record<string, never>).ack).status;
        if (status !== 'applied') failures.push(`${label}: ${step} ${status}`);
        groups.push(sent);
      }
      const late = groups.at(-2)!, early = groups.at(-1)!;
      for (const update of groups.slice(0, -2).flat()) await b.session.submit('applyRemoteUpdate', { update }, merge).ack;
      const vector = b.engine.encodeStateVector().slice();
      for (const update of [...early, ...late]) {
        const outcome = await b.session.submit('applyRemoteUpdate', { update }, merge).ack.then(ack => ack.status, (error: unknown) => String(error));
        if (outcome !== 'applied' && outcome !== 'noop') {
          const held = early.includes(update) && Buffer.compare(vector, b.engine.encodeStateVector()) === 0;
          failures.push(`${label}: ${outcome} (structs held after the early update: ${held}, failed: ${b.client.hasFailed()})`);
          break;
        }
        if (update === early.at(-1) && b.engine.hasPendingUpdate()) awaited.push(label);
      }
      if (b.client.hasFailed()) continue;
      for (const update of back) await a.session.submit('applyRemoteUpdate', { update }, merge).ack;
      expect(clocks(b.engine.encodeStateVector())).toEqual(clocks(a.engine.encodeStateVector()));
      expect(b.engine.paragraphs('body').map(p => p.text)).toEqual(a.engine.paragraphs('body').map(p => p.text));
      expect(b.engine.hasPendingUpdate()).toBe(false);
      const at = { story: 'body', paraId: b.engine.paragraphs('body')[0]!.paraId, offset: 0 };
      expect((await b.session.submit('insertText', { selection: { anchor: at, focus: at }, text: 'x' }).ack).status).toBe('applied');
    } finally { offA(); offB(); a.session.dispose(); b.session.dispose(); }
  }
  expect(failures).toEqual([]);
  expect(awaited.length).toBeGreaterThan(0);
}, TIMEOUT * 3);

test('a worker replica whose projection fails with no update pending fails loudly', async () => {
  const a = await open(9201);
  const b = await open(9202, a.engine.encodeState());
  const errors: CollaborationError[] = [];
  let receive: ((event: CollaborationTransportEvent) => void) | undefined;
  let applying: Promise<unknown> | undefined;
  const transport: CollaborationTransport = {
    connect() {}, disconnect() {}, send() { return true; },
    onEvent(listener) { receive = listener; return () => { receive = undefined; }; },
  };
  const provider = new CollaborationProvider(createAsyncCollaborationReplica({
    clientId: 9202,
    encodeStateVector: async () => b.engine.encodeStateVector(),
    encodeStateAsUpdate: async vector => b.engine.encodeStateAsUpdate(vector),
    applyUpdate: update => (applying = b.session.submit('applyRemoteUpdate', { update: update.slice().buffer }, merge).ack),
    onUpdate: listener => b.client.onCollabUpdates(event => event.updates.forEach(update => listener(update, 'local'))),
  }), transport);
  provider.onError(error => errors.push(error));
  const pending: boolean[] = [];
  spyOn(b.engine, 'layoutDocumentWithRegionsAfterEdit').mockImplementation(() => {
    pending.push(b.engine.hasPendingUpdate());
    throw new Error('story "body" contains an unsupported embed at 3');
  });
  try {
    provider.connect();
    receive!({ type: 'open' });
    expect((await a.session.submit('insertText', { text: '!' }).ack).status).toBe('applied');
    receive!({ type: 'message', data: encodeUpdate(new Uint8Array(remoteUpdate(b, a.engine))) });
    await applying?.catch(() => {});
    await new Promise(resolve => setTimeout(resolve, 0));
    expect(pending.length).toBeGreaterThan(0);
    expect(pending).not.toContain(true);
    expect(errors[0]?.message).toContain('unsupported embed at 3');
    expect(b.client.hasFailed()).toBe(true);
  } finally { provider.destroy(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

test('a comment that reaches a worker replica before the text it is anchored on lays out and resolves once the text arrives', async () => {
  const a = await open(9201);
  const b = await open(9202, a.engine.encodeState());
  const c = await open(9203, a.engine.encodeState());
  let fromA: ArrayBuffer[] = [];
  const fromB: ArrayBuffer[] = [];
  const offA = a.client.onCollabUpdates(event => fromA.push(...event.updates.map(u => u.slice().buffer)));
  const offB = b.client.onCollabUpdates(event => fromB.push(...event.updates.map(u => u.slice().buffer)));
  const apply = (target: typeof c, update: ArrayBuffer) =>
    target.session.submit('applyRemoteUpdate', { update }, merge).ack.then(ack => ack.status, (error: unknown) => String(error));
  try {
    expect((await b.session.submit('insertText', { text: 'B' }).ack).status).toBe('applied');
    for (const update of fromB) expect(await apply(a, update)).toBe('applied');
    fromA = [];
    const [first, second] = a.engine.paragraphs('body').map(paragraph => paragraph.paraId);
    const range = { story: 'body', start: { paraId: first!, offset: 4 }, end: { paraId: first!, offset: 5 } };
    expect((await a.session.submit('addComment', { ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Note',
      expectVersion: a.engine.proposalEngine.version() }).ack).status).toBe('applied');
    const at = { story: 'body', paraId: second!, offset: 0 };
    expect((await a.session.submit('insertText', { selection: { anchor: at, focus: at }, text: 'A' }).ack).status).toBe('applied');
    expect(fromA.length).toBeGreaterThan(0);

    for (const update of fromA) expect(await apply(c, update)).toBe('applied');
    expect(c.client.hasFailed()).toBe(false);
    const id = a.engine.listComments()[0]!.id;
    expect(c.engine.listComments().map(comment => comment.id)).toEqual([id]);
    expect(await commentedText(c)).toBe('');
    expect((await c.client.save({ comments: [] })).bytes.byteLength).toBeGreaterThan(0);
    for (const update of fromB) expect(await apply(c, update)).toBe('applied');
    expect(c.engine.resolveComment(id)).toEqual(a.engine.resolveComment(id));
    const highlighted = await commentedText(a);
    expect(highlighted).not.toBe('');
    expect(await commentedText(c)).toBe(highlighted);
    expect(c.engine.paragraphs('body')).toEqual(a.engine.paragraphs('body'));
    const end = { story: 'body', paraId: first!, offset: 5 };
    expect((await c.session.submit('insertText', { selection: { anchor: end, focus: end }, text: 'C' }).ack).status).toBe('applied');
  } finally { offA(); offB(); a.session.dispose(); b.session.dispose(); c.session.dispose(); }
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
    const ack = await a.client.edit({ session: a.initial.session, incarnation: 1, seq: 1,
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

test('a byte source saves the shared stories it joined with and leaves a plain open untouched', async () => {
  const bytes = documentBytes('HeaderSeed');
  const remote = await createYrsSession({ clientId: 9203 });
  const plain = await open(9201, undefined, undefined, bytes);
  const saveAfterTyping = async (opened: Awaited<ReturnType<typeof open>>) => {
    expect((await opened.session.submit('insertText', { text: 'B' }).ack).status).toBe('applied');
    const saved = new Uint8Array((await opened.client.save({ comments: [] })).bytes);
    const reopened = await createYrsSession({ clientId: 9298 });
    try {
      reopened.openDocx(saved.slice(), true);
      return { saved, body: reopened.paragraphs('body')[0]!.text, header: reopened.paragraphs('hf:rIdH1').map(paragraph => paragraph.text) };
    } finally { reopened.destroy(); }
  };
  try {
    const untouched = await saveAfterTyping(plain);
    expect(untouched.body).toBe('SeedB');
    expect(untouched.header).toEqual(['HeaderSeed']);
    expect(unzipContainer(untouched.saved)['word/header1.xml']).toEqual(unzipContainer(bytes)['word/header1.xml']);

    remote.openDocx(bytes.slice(), true);
    remote.insertText({ story: 'hf:rIdH1', paraId: remote.paragraphs('hf:rIdH1')[0]!.paraId, offset: 10 }, ' shared');
    const shared = await open(9202, remote.encodeState(), undefined, bytes);
    try {
      const joined = await saveAfterTyping(shared);
      expect(joined.body).toBe('SeedB');
      expect(joined.header).toEqual(['HeaderSeed shared']);
    } finally { shared.session.dispose(); }
  } finally { remote.destroy(); plain.session.dispose(); }
}, TIMEOUT);

test('worker replicas that number comments in different local orders save and reopen the same comments', async () => {
  const a = await open(9202);
  const b = await open(9201, a.engine.encodeState());
  const comment = (replica: Replica, index: number, body: string) => {
    const paraId = replica.engine.paragraphs('body')[index]!.paraId;
    return replica.session.submit('addComment', { ranges: [{ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 3 } }],
      author: 'Ada', date: '2026-01-01', body, expectVersion: replica.engine.proposalEngine.version() }).ack.then(ack => ack.status);
  };
  const exchange = (target: Replica, source: Replica) =>
    target.session.submit('applyRemoteUpdate', { update: remoteUpdate(target, source.engine) }, merge).ack.then(ack => ack.status);
  const saved = async (replica: Replica) => {
    const bytes = new Uint8Array((await replica.client.save({ comments: [] })).bytes);
    const parts = unzipContainer(bytes);
    const xml = new TextDecoder().decode(parts['word/comments.xml']!);
    const ids = Object.fromEntries([...xml.matchAll(/<w:comment w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:comment>/gu)]
      .map(match => [[...match[2]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/gu)].map(run => run[1]).join(''), match[1]]));
    const markers = new TextDecoder().decode(parts['word/document.xml']!).split('</w:p>').slice(0, 2)
      .map(paragraph => [...paragraph.matchAll(/<w:comment(?:RangeStart|RangeEnd|Reference) w:id="(\d+)"/gu)].map(match => match[1]));
    const reopened = await createYrsSession({ clientId: 9298 });
    try {
      reopened.openDocx(bytes.slice(), true);
      return { xml, ids, markers, reopened: reopened.listComments().map(listed => listed.id).sort() };
    } finally { reopened.destroy(); }
  };
  try {
    expect(await comment(a, 0, 'First')).toBe('applied');
    expect((await saved(a)).ids).toEqual({ First: '1' });
    expect(await comment(b, 1, 'Second')).toBe('applied');
    expect((await saved(b)).ids).toEqual({ Second: '1' });
    expect(await exchange(a, b)).toBe('applied');
    expect(await exchange(b, a)).toBe('applied');
    const left = await saved(a);
    expect(Object.values(left.ids).sort()).toEqual(['1', '2']);
    expect(left.markers).toEqual([Array(3).fill(left.ids.First), Array(3).fill(left.ids.Second)]);
    expect(left.reopened).toEqual(['1', '2']);
    expect(await saved(b)).toEqual(left);
  } finally { a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

test('a merge that deletes the story holding the local selection repairs the confirmed selection', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    const table = remote.insertTable({ story: 'body', paraId: a.engine.paragraphs('body')[0]!.paraId, offset: 0 }, 1, 1);
    const cell = table.createdStoryIds[0]!;
    expect((await a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge).ack).status).toBe('applied');
    const inCell = { story: cell, paraId: a.engine.paragraphs(cell)[0]!.paraId, offset: 0 };
    await a.session.submit('select', { selection: { anchor: inCell, focus: inCell } }).ack;
    expect(a.session.state.selection?.focus.story).toBe(cell);
    remote.deleteTable(table.table);
    expect((await a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge).ack).status).toBe('applied');
    const repaired = a.session.state.selection;
    expect(repaired?.focus.story).toBe('body');
    expect(a.engine.paragraphs('body').map(paragraph => paragraph.paraId)).toContain(repaired!.focus.paraId);
    expect((await a.session.submit('insertText', { text: 'Z' }).ack).status).toBe('applied');
    expect(a.engine.paragraphs('body').some(paragraph => paragraph.text.includes('Z'))).toBe(true);
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('a merge that deletes header content repairs the local selection within the header story', async () => {
  const a = await open(9201, undefined, undefined, documentBytes('HeaderSeed'));
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    const story = 'hf:rIdH1';
    const paraId = a.engine.paragraphs(story)[0]!.paraId;
    const inHeader = { story, paraId, offset: 6 };
    expect((await a.session.submit('select', { selection: { anchor: inHeader, focus: inHeader } }).ack).status).toBe('noop');
    expect(a.session.state.selection?.focus).toMatchObject(inHeader);
    remote.deleteRange({ story, start: { paraId, offset: 0 }, end: { paraId, offset: 10 } });
    expect((await a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge).ack).status).toBe('applied');
    expect(a.engine.paragraphs(story).map(paragraph => paragraph.text)).toEqual(['']);
    const repaired = { story, paraId, offset: 0 };
    expect(a.session.state.selection).toMatchObject({ anchor: repaired, focus: repaired });
    expect((await a.session.submit('insertText', { text: 'Z' }).ack).status).toBe('applied');
    expect(a.engine.paragraphs(story).map(paragraph => paragraph.text)).toEqual(['Z']);
    expect(a.engine.paragraphs('body').map(paragraph => paragraph.text)).toEqual(['Seed', 'Second']);
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('ordinary typing behind an unacknowledged merge resolves through sticky positions', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    remote.insertText({ story: 'body', paraId: a.engine.paragraphs('body')[0]!.paraId, offset: 0 }, 'R');
    const merged = a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge);
    const typing = a.session.submit('insertText', { text: '!' });
    expect((await merged.ack).status).toBe('applied');
    expect((await typing.ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toBe('RSeed!');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('ordinary typing behind a merge that deletes its selection story is refused', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    const table = remote.insertTable({ story: 'body', paraId: a.engine.paragraphs('body')[0]!.paraId, offset: 0 }, 1, 1);
    await a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge).ack;
    const story = table.createdStoryIds[0]!;
    const at = { story, paraId: a.engine.paragraphs(story)[0]!.paraId, offset: 0 };
    await a.session.submit('select', { selection: { anchor: at, focus: at } }).ack;
    remote.deleteTable(table.table);
    const merged = a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge);
    const typing = a.session.submit('insertText', { text: 'misplaced' });
    expect((await merged.ack).status).toBe('applied');
    expect(await typing.ack).toMatchObject({ status: 'refused', reason: 'stale-anchor' });
    expect(a.engine.paragraphs('body').map(p => p.text)).toEqual(['Seed', 'Second']);
    expect((await a.session.submit('insertText', { text: 'fresh' }).ack).status).toBe('applied');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('input behind an unacknowledged local edit or expecting its version is still refused across a merge', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    const paraId = a.engine.paragraphs('body')[0]!.paraId;
    remote.insertText({ story: 'body', paraId, offset: 0 }, 'R');
    const local = a.session.submit('insertText', { text: 'a' });
    const merged = a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge);
    const behindLocal = a.session.submit('insertText', { text: 'b' });
    expect((await local.ack).status).toBe('applied');
    expect((await merged.ack).status).toBe('applied');
    expect(await behindLocal.ack).toMatchObject({ status: 'refused', reason: 'stale-anchor' });
    remote.insertText({ story: 'body', paraId, offset: 0 }, 'Q');
    const again = a.session.submit('applyRemoteUpdate', { update: remoteUpdate(a, remote) }, merge);
    const expecting = a.session.submit('insertText', { text: 'c' }, { expectVersion: true });
    expect((await again.ack).status).toBe('applied');
    expect(await expecting.ack).toMatchObject({ status: 'refused', reason: 'stale-version' });
    expect(a.engine.paragraphs('body')[0]!.text).toBe('QRSeeda');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('a remote update past the relay frame cap applies and one past the engine cap is refused', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    remote.setParagraphAttr(a.engine.paragraphs('body')[0]!.paraId, 'bulk', 'x'.repeat(20 * 1024 * 1024));
    const update = remoteUpdate(a, remote);
    expect(update.byteLength).toBeGreaterThan(MAX_COLLABORATION_FRAME_BYTES);
    expect((await a.session.submit('applyRemoteUpdate', { update }, merge).ack).status).toBe('applied');
    expect(await a.session.submit('applyRemoteUpdate', { update: new ArrayBuffer(MAX_UPDATE_BYTES + 1) }, merge).ack)
      .toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    expect((await a.session.submit('insertText', { text: '!' }).ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed!');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('chunked remote updates merge as one change and malformed or oversized chunk sets are refused', async () => {
  const a = await open(9201);
  const remote = await createYrsSession({ clientId: 9203 });
  try {
    remote.loadState(a.engine.encodeState());
    const paraId = a.engine.paragraphs('body')[0]!.paraId;
    for (const word of ['one', 'two', 'three', 'four', 'five', 'six']) remote.insertText({ story: 'body', paraId, offset: 0 }, `${word} `);
    const chunks = remote.encodeStateAsUpdateChunks(a.engine.encodeStateVector(), 24).map(chunk => chunk.slice().buffer);
    expect(chunks.length).toBeGreaterThan(1);
    expect((await a.session.submit('applyRemoteUpdateChunks', { updates: chunks }, merge).ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toBe('six five four three two one Seed');
    expect(a.engine.encodeStateVector()).toEqual(remote.encodeStateVector());
    const state = a.engine.encodeState();
    const large = new ArrayBuffer(MAX_UPDATE_BYTES);
    for (const updates of [[], [new ArrayBuffer(0)], [new ArrayBuffer(MAX_UPDATE_BYTES + 1)], [chunks[0]!, new ArrayBuffer(0)],
      [...Array.from({ length: MAX_SYNC_CHUNKS_BYTES / MAX_UPDATE_BYTES }, () => large), new ArrayBuffer(1)]]) {
      expect(await a.session.submit('applyRemoteUpdateChunks', { updates }, merge).ack).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    }
    expect(a.engine.encodeState()).toEqual(state);
    expect((await a.session.submit('insertText', { text: '!' }).ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toBe('six five four three two one Seed!');
  } finally { remote.destroy(); a.session.dispose(); }
}, TIMEOUT);

test('provider decode rejects an oversized inbound frame and local editing still applies', async () => {
  const a = await open(9201);
  const errors: CollaborationError[] = [];
  let receive: ((event: CollaborationTransportEvent) => void) | undefined;
  const transport: CollaborationTransport = {
    connect() {}, disconnect() {}, send() { return true; },
    onEvent(listener) { receive = listener; return () => { receive = undefined; }; },
  };
  const provider = new CollaborationProvider(createAsyncCollaborationReplica({
    clientId: 9201,
    encodeStateVector: async () => a.engine.encodeStateVector(),
    encodeStateAsUpdate: async vector => a.engine.encodeStateAsUpdate(vector),
    applyUpdate: update => a.session.submit('applyRemoteUpdate', { update: update.slice().buffer }, merge).ack,
    onUpdate: listener => a.client.onCollabUpdates(event => event.updates.forEach(update => listener(update, 'local'))),
  }), transport);
  provider.onError(error => errors.push(error));
  try {
    provider.connect();
    receive!({ type: 'open' });
    const before = a.engine.encodeState();
    const edits = a.worker.requests.filter(kind => kind === 'edit').length;
    receive!({ type: 'message', data: new Uint8Array(MAX_COLLABORATION_FRAME_BYTES + 1) });
    expect(errors).toHaveLength(1);
    expect(errors[0]).toMatchObject({ code: 'protocol' });
    expect(errors[0]!.message).toContain('Frame exceeds');
    expect(a.worker.requests.filter(kind => kind === 'edit')).toHaveLength(edits);
    expect(a.engine.encodeState()).toEqual(before);
    expect((await a.session.submit('insertText', { text: '!' }).ack).status).toBe('applied');
    expect(a.engine.paragraphs('body')[0]!.text).toBe('Seed!');
  } finally { provider.destroy(); a.session.dispose(); }
}, TIMEOUT);

type Replica = Awaited<ReturnType<typeof open>>;

function relay(from: Replica, to: Replica) {
  const pending: ArrayBuffer[] = [];
  const off = from.client.onCollabUpdates(event => pending.push(...event.updates.map(update => update.slice().buffer)));
  return { off, async flush() { for (const update of pending.splice(0)) await to.session.submit('applyRemoteUpdate', { update }, merge).ack; } };
}

const pageTexts = (frame: RetainedFrame) => frame.displayList.pages.map(({ primitives }) => primitives
  .map(primitive => primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : '').join(''));

const commentedText = async (replica: Replica) => applyFrameDeltaOwned(null, decodeFrameDelta((await replica.client.requestFullFrame()).frame!.payload.frame))
  .displayList.pages.flatMap(({ primitives }) => primitives)
  .map(primitive => (primitive.kind === 'glyphRun' || primitive.kind === 'text') && primitive.commentIds?.length ? primitive.text : '').join('');

async function sectionView(replica: Replica) {
  const frame = applyFrameDeltaOwned(null, decodeFrameDelta((await replica.client.requestFullFrame()).frame!.payload.frame));
  const xml = new TextDecoder().decode(unzipContainer(new Uint8Array((await replica.client.save({ comments: [] })).bytes))['word/document.xml']!);
  return { pages: pageTexts(frame), sections: replica.engine.sectionProperties().sections.map(({ properties }) => properties.sectionStart ?? null),
    body: xml.slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>')) };
}

const projectedSections = (replica: Replica) => replica.engine.partDocument().package.document.sections?.map(section => section.properties);

const SMALL_PAGE = '<w:pgSz w:w="6000" w:h="4000"/><w:pgMar w:top="500" w:right="500" w:bottom="500" w:left="500" w:header="200" w:footer="200" w:gutter="0"/>';

test('a merged section insert lays out and saves like the local edit with a header', async () => {
  const bytes = documentBytes('HeaderSeed');
  const a = await open(9201, undefined, undefined, bytes);
  const b = await open(9202, a.engine.encodeState(), undefined, bytes);
  const toA = relay(b, a);
  try {
    const before = await sectionView(b);
    const at = { story: 'body', paraId: b.engine.paragraphs('body')[1]!.paraId, offset: 3 };
    expect((await b.session.submit('insertSectionBreak', { type: 'nextPage', selection: { anchor: at, focus: at } }).ack).status).toBe('applied');
    await toA.flush();
    const inserted = await sectionView(b);
    expect(inserted.sections).toHaveLength(before.sections.length + 1);
    expect(inserted.pages.length).toBeGreaterThan(before.pages.length);
    expect(await sectionView(a)).toEqual(inserted);
    expect(projectedSections(a)).toEqual(projectedSections(b));
  } finally { toA.off(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

test('a merged section delete lays out and saves like the local edit after joining', async () => {
  const bytes = documentBytes('HeaderSeed');
  const a = await open(9201, undefined, undefined, bytes);
  let b: Replica | undefined;
  try {
    const at = { story: 'body', paraId: a.engine.paragraphs('body')[1]!.paraId, offset: 3 };
    expect((await a.session.submit('insertSectionBreak', { type: 'nextPage', selection: { anchor: at, focus: at } }).ack).status).toBe('applied');
    b = await open(9202, a.engine.encodeState(), undefined, bytes);
    const before = await sectionView(a);
    const [seed, closing] = a.engine.paragraphs('body');
    const selection = { anchor: { story: 'body', paraId: seed!.paraId, offset: 2 }, focus: { story: 'body', paraId: closing!.paraId, offset: 1 } };
    expect((await a.session.submit('deleteBackward', { selection }).ack).status).toBe('applied');
    expect((await b.session.submit('applyRemoteUpdate', { update: remoteUpdate(b, a.engine) }, merge).ack).status).toBe('applied');
    const deleted = await sectionView(a);
    expect(deleted.sections).toHaveLength(before.sections.length - 1);
    expect(await sectionView(b)).toEqual(deleted);
    expect(projectedSections(b)).toEqual(projectedSections(a));
  } finally { a.session.dispose(); b?.session.dispose(); }
}, TIMEOUT);

test.todo('a local delete across a seeded section boundary must mark sectionsTouched for save parity with a merge', async () => {
  const bytes = documentBytes(undefined, `<w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>${SMALL_PAGE}</w:sectPr></w:pPr>` +
    '<w:r><w:t>Closing</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>');
  const a = await open(9201, undefined, undefined, bytes);
  const b = await open(9202, a.engine.encodeState(), undefined, bytes);
  const toA = relay(b, a);
  try {
    const before = await sectionView(b);
    expect(await sectionView(a)).toEqual(before);
    const [seed, closing] = b.engine.paragraphs('body');
    const selection = { anchor: { story: 'body', paraId: seed!.paraId, offset: 2 }, focus: { story: 'body', paraId: closing!.paraId, offset: 2 } };
    expect((await b.session.submit('deleteBackward', { selection }).ack).status).toBe('applied');
    await toA.flush();
    const deleted = await sectionView(b);
    expect(deleted.sections).toHaveLength(before.sections.length - 1);
    expect(await sectionView(a)).toEqual(deleted);
  } finally { toA.off(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

const proposalIn = (replica: Replica): DocxProposalInput => ({ id: 'proposal', paragraph: replica.engine.paragraphIdentities().paragraphs[0]!.session!,
  suggest: { author: 'Host', date: '2026-10-13T00:00:00Z' }, op: 'replaceText', search: 'e', replaceWith: 'E', occurrence: 'all' });
const pendingRevisions = (replica: Replica) => replica.engine.proposalEngine.listRevisions().map(({ revisionId }) => revisionId);

async function proposalsOf(replica: Replica) {
  await replica.client.requestFullFrame();
  return (await replica.client.documentRead({ kind: 'reviewSnapshot' })).value!.proposals.proposals
    .map(({ id, revisionIds }) => ({ id, revisionIds: [...revisionIds] }));
}

for (const settle of ['all', 'some', 'withdraw'] as const) {
  test(`a proposal made in a room reaches the peer and ${settle === 'withdraw' ? 'its withdrawal clears both replicas' : `the peer accepting ${settle} of it narrows the card`}`, async () => {
    const bytes = documentBytes(undefined, '<w:p><w:r><w:t>Seed and seed</w:t></w:r></w:p>');
    const a = await open(9201, undefined, undefined, bytes);
    const b = await open(9202, a.engine.encodeState(), undefined, bytes);
    const toA = relay(b, a), toB = relay(a, b);
    try {
      const proposed = await a.session.submit('propose', { expectVersion: a.engine.proposalEngine.version(), proposals: [proposalIn(a)] }).ack;
      expect(proposed).toMatchObject({ status: 'applied' });
      const [card] = await proposalsOf(a);
      expect(card!.revisionIds.length).toBeGreaterThan(1);
      await toB.flush();
      expect(pendingRevisions(b)).toEqual(expect.arrayContaining(card!.revisionIds));
      if (settle === 'withdraw') {
        expect((await a.session.submit('withdrawProposals', { ids: ['proposal'], expectVersion: a.engine.proposalEngine.version() }).ack).status).toBe('applied');
        await toB.flush();
        expect(pendingRevisions(b).filter(id => card!.revisionIds.includes(id))).toEqual([]);
        expect(await proposalsOf(a)).toEqual([]);
        expect(await proposalsOf(b)).toEqual([]);
        expect(b.engine.paragraphs('body')).toEqual(a.engine.paragraphs('body'));
        return;
      }
      for (const revisionId of settle === 'all' ? card!.revisionIds : card!.revisionIds.slice(0, 1)) {
        if (!pendingRevisions(b).includes(revisionId)) continue;
        expect((await b.session.submit('acceptRevision', { revisionId, expectVersion: b.engine.proposalEngine.version() }).ack).status).toBe('applied');
      }
      await toA.flush();
      const remaining = card!.revisionIds.filter(id => pendingRevisions(a).includes(id));
      if (settle === 'all') expect(remaining).toEqual([]);
      else expect(remaining.length).toBeGreaterThan(0);
      expect(remaining.length).toBeLessThan(card!.revisionIds.length);
      expect(await proposalsOf(a)).toEqual(remaining.length ? [{ id: 'proposal', revisionIds: remaining }] : []);
      expect(await proposalsOf(b)).toEqual([]);
      if (settle === 'all') {
        const accepted = a.engine.paragraphs('body');
        expect((await a.session.submit('withdrawProposals', { ids: ['proposal'], expectVersion: a.engine.proposalEngine.version() }).ack).status).toBe('noop');
        await toB.flush();
        expect(a.engine.paragraphs('body')).toEqual(accepted);
        expect(b.engine.paragraphs('body')).toEqual(accepted);
        expect(await proposalsOf(a)).toEqual([]);
        expect(await proposalsOf(b)).toEqual([]);
      }
    } finally { toA.off(); toB.off(); a.session.dispose(); b.session.dispose(); }
  }, TIMEOUT);
}

test('a merge deleting proposed text removes the proposal card on both replicas', async () => {
  const a = await open(9201);
  const b = await open(9202, a.engine.encodeState());
  const remote = await createYrsSession({ clientId: 9203 });
  const toA = relay(b, a), toB = relay(a, b);
  try {
    const second = b.engine.paragraphIdentities().paragraphs[1]!.session!;
    for (const [replica, paragraph] of [[a, proposalIn(a).paragraph], [b, second]] as const) {
      expect((await replica.session.submit('propose', { expectVersion: replica.engine.proposalEngine.version(),
        proposals: [{ ...proposalIn(replica), paragraph }] }).ack).status).toBe('applied');
    }
    await toB.flush();
    await toA.flush();
    for (const replica of [a, b]) {
      const [card] = await proposalsOf(replica);
      expect(card!.revisionIds.length).toBeGreaterThan(0);
    }
    remote.loadState(a.engine.encodeState());
    const first = remote.paragraphs('body')[0]!;
    const last = remote.paragraphs('body').at(-1)!;
    remote.deleteRange({ story: 'body', start: { paraId: first.paraId, offset: 0 }, end: { paraId: last.paraId, offset: last.text.length } });
    for (const replica of [a, b]) {
      expect((await replica.session.submit('applyRemoteUpdate', { update: remoteUpdate(replica, remote) }, merge).ack).status).toBe('applied');
      expect(await proposalsOf(replica)).toEqual([]);
      expect(pendingRevisions(replica)).toEqual([]);
    }
    expect(a.engine.paragraphs('body')).toEqual(b.engine.paragraphs('body'));
  } finally { toA.off(); toB.off(); remote.destroy(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

test('two opens that both seed from one source merge by story last-writer-wins like the in-thread route', async () => {
  const a = await open(9201);
  const b = await open(9202);
  const left = await createYrsSession({ clientId: 9201 });
  const right = await createYrsSession({ clientId: 9202 });
  const texts = (session: { paragraphs(story: string): readonly { text: string }[] }) => session.paragraphs('body').map(paragraph => paragraph.text);
  try {
    left.openDocx(documentBytes(), true);
    right.openDocx(documentBytes(), true);
    left.insertText({ story: 'body', paraId: left.paragraphs('body')[0]!.paraId, offset: 4 }, 'A');
    right.insertText({ story: 'body', paraId: right.paragraphs('body')[0]!.paraId, offset: 4 }, 'B');
    const [toLeft, toRight] = [right.encodeStateAsUpdate(left.encodeStateVector()), left.encodeStateAsUpdate(right.encodeStateVector())];
    left.applyUpdate(toLeft);
    right.applyUpdate(toRight);
    expect(texts(left)).toEqual(texts(right));
    expect([['SeedA', 'Second'], ['SeedB', 'Second']]).toContainEqual(texts(left));

    expect((await a.session.submit('insertText', { text: 'A' }).ack).status).toBe('applied');
    expect((await b.session.submit('insertText', { text: 'B' }).ack).status).toBe('applied');
    const [toA, toB] = [remoteUpdate(a, b.engine), remoteUpdate(b, a.engine)];
    await a.session.submit('applyRemoteUpdate', { update: toA }, merge).ack;
    await b.session.submit('applyRemoteUpdate', { update: toB }, merge).ack;
    expect(texts(a.engine)).toEqual(texts(b.engine));
    expect(texts(a.engine)).toEqual(texts(left));
  } finally { left.destroy(); right.destroy(); a.session.dispose(); b.session.dispose(); }
}, TIMEOUT);

const presenceRead = (replica: Replica, expectVersion = replica.engine.proposalEngine.version()) => replica.client.documentRead({ kind: 'collabPresence', pages: [0, 1], expectVersion,
  expectFrameEpoch: replica.initial.frame!.revision, expectLayoutRevision: replica.initial.frame!.payload.layoutRevision, peers: [] });
const die = (replica: Replica) => replica.worker.onerror?.({ message: 'injected worker death' } as ErrorEvent);
const outcomes = async (pending: Promise<unknown>[]) => (await Promise.allSettled(pending))
  .map(result => result.status === 'rejected' ? 'rejected' : [(result.value as { status?: string }).status ?? 'fulfilled', (result.value as { reason?: string }).reason].filter(Boolean).join(':'));

test('worker death during a collaboration read settles the read and every pending ticket with a failure', async () => {
  const a = await open(9201);
  try {
    const before = a.engine.encodeState();
    const version = a.engine.proposalEngine.version();
    let atDeath: Uint8Array | undefined;
    const destroy = a.engine.destroy;
    a.engine.destroy = () => { atDeath ??= a.engine.encodeState(); destroy(); };
    const pending = [presenceRead(a), a.session.submit('insertText', { text: '!' }).ack];
    die(a);
    expect(await outcomes(pending)).toEqual(['rejected', 'rejected']);
    expect(atDeath).toEqual(before);
    expect(await outcomes([presenceRead(a, version), a.client.documentRead({ kind: 'collabStateVector' })])).toEqual(['rejected', 'rejected']);
  } finally { a.session.dispose(); }
}, TIMEOUT);

for (const failure of ['a thrown merge', 'worker death in a merge'] as const) {
  test(`${failure} settles every pending ticket and later collaboration read with a failure`, async () => {
    const create = wasm.createEditSession;
    let raw: ReturnType<typeof create> | undefined;
    const capture = spyOn(wasm, 'createEditSession').mockImplementation(clientId => {
      const session = create(clientId);
      if (clientId === 9201) raw = session;
      return session;
    });
    let a: Replica;
    try { a = await open(9201); } finally { capture.mockRestore(); }
    try {
      const version = a.engine.proposalEngine.version();
      const before = a.engine.encodeState();
      let stateAtFailure: Uint8Array | undefined;
      const destroy = a.engine.destroy;
      a.engine.destroy = () => {
        stateAtFailure ??= a.engine.encodeState();
        destroy();
      };
      const payloadText = 'PRIVATE-UPDATE-CONTENT';
      const payload = new TextEncoder().encode(payloadText);
      const apply = raw!.apply_update_with_inference.bind(raw!);
      let entered = false;
      raw!.apply_update_with_inference = update => {
        entered = true;
        try { return apply(update); }
        catch {
          if (failure === 'worker death in a merge') die(a);
          throw new WebAssembly.RuntimeError(payloadText);
        }
      };
      const pending = [a.session.submit('applyRemoteUpdate', { update: payload.buffer }, merge).ack,
        a.session.submit('insertText', { text: '!' }).ack];
      const settled = await Promise.allSettled(pending);
      expect(await outcomes(pending)).toEqual([
          failure === 'a thrown merge' ? 'refused:unknown-outcome' : 'rejected', 'rejected',
        ]);
      expect(entered).toBe(true);
      const messages = settled.map(result => result.status === 'rejected' ? String(result.reason) : JSON.stringify(result.value)).join('\n');
      expect(messages).not.toContain(payloadText);
      expect(messages).not.toContain([...payload].join(','));
      expect(stateAtFailure ?? a.engine.encodeState()).toEqual(before);
      expect(await outcomes([presenceRead(a, version)])).toEqual(['rejected']);
      expect(await outcomes([a.session.submit('insertText', { text: 'later' }).ack])).toEqual(['rejected']);
      expect(await outcomes([a.session.submit('insertText', { text: 'again' }).ack])).toEqual(['rejected']);
    } finally { a.session.dispose(); }
  }, TIMEOUT);
}

import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { createResidentCommandConnection } from './residentCommandConnection';
import { createDocxWorkerEditSession } from './createDocxWorkerEditSession';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';

let start: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => { await openCommandEngine(); start = await residentWorkerFactory(); });
afterEach(closeCommandEngines);

function documentBytes() {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')],
  ])));
}

async function open(clientId: number, initialUpdate?: Uint8Array) {
  const worker = start(clientId);
  const client = new ResidentEngineWorkerClient(worker);
  await client.open(documentBytes(), { collaboration: { clientId, initialUpdate } });
  const engine = worker.sessions[0]!;
  const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 4 };
  const initial = await client.commandInit({ session: `replica-${clientId}`, incarnation: 1, epoch: 1,
    readOnly: false, selection: { anchor: at, focus: at },
    fonts: [new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))],
    layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} }),
    layoutWithRegions: true, layoutRevision: 1, extras: '{}',
  });
  const session = createDocxWorkerEditSession({ source: documentBytes(),
    connection: await createResidentCommandConnection(client, initial),
    echo: { paint() {}, drop() {}, present() {} }, maxRecoveryAttempts: 0,
    recovery: { reopen: () => Promise.reject(new Error('Unexpected reopen')) },
  });
  return { worker, client, engine, initial, session };
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
  const off = a.client.onCollabUpdates(() => { events.push('update'); });
  try {
    await a.session.submit('insertText', { text: 'A' }).ack.then(() => events.push('ack'));
    expect(events).toEqual(['update', 'ack']);
    await b.session.submit('insertText', { text: 'B' }).ack;
    const merge = a.session.submit('applyRemoteUpdate', { update: b.engine.encodeStateAsUpdate(a.engine.encodeStateVector()).buffer }, { lane: 'collab', fence: true });
    const typing = a.session.submit('insertText', { text: '!' }, { afterSeq: merge.seq });
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
  try {
    const cursor = a.engine.encodeSelection()!;
    const version = a.engine.proposalEngine.version();
    const read = { kind: 'collabPresence' as const, peers: [{ clientId: 9202, clock: 1, cursor }], pages: [0, 1] as [number, number],
      expectVersion: version, expectFrameEpoch: a.initial.frame.revision, expectLayoutRevision: 1 };
    expect((await a.client.documentRead(read)).value?.peers[0]).toMatchObject({ status: 'ok', caret: { pageIndex: 0, pageId: '1' } });
    expect((await a.client.documentRead({ ...read, expectFrameEpoch: 999 })).value).toBeNull();
    expect((await a.client.documentRead({ ...read, peers: [{ clientId: 9202, clock: 2,
      cursor: { ...cursor, anchor: new Uint8Array([255]) } }] })).value?.peers[0]?.status).toBe('unresolved');
  } finally { a.session.dispose(); }
}, TIMEOUT);

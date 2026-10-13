import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/yrs/wasm/index';
import { preloadOpcWasm } from '@betteroffice/docx/wasm/opc';
import { residentWorkerFactory } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { ResidentSourceWorkerClient, createDocxWorkerEditSession, createResidentCommandConnection } from '@betteroffice/docx/peerless';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';

const TIMEOUT = 20_000;
const generated = resolve(import.meta.dir, '../../../../../docx/src/wasm/generated');
let start: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(generated, 'edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(generated, 'opc/ooxml_opc_bg.wasm'))));
  start = await residentWorkerFactory();
});

const documentBytes = () => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
  ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')],
])));

async function open(clientId: number) {
  const worker = start(clientId);
  const client = new ResidentSourceWorkerClient(worker);
  await client.openSource({ bytes: documentBytes(), collaboration: { clientId } });
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
  const replica = createPeerlessCollaborationReplica(client, session, clientId);
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

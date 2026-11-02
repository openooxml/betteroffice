import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { initWasm } from '../wasm/loader';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from './commandClient';
import { createPresentationSessionHost } from './host';
import { commandMetafileSvgModule, type MetafileSvgModule } from './metafileSvg';
import type { PresentationCommandClient } from './commandProtocol';

const clients: PresentationCommandClient[] = [];
const restorers: (() => void)[] = [];
const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96"/>');

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function deck(pictureSlide?: number, hiddenFromScan = false): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="bin" ContentType="image/x-emf"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${[1, 2].map((number) => `<Override PartName="/ppt/slides/slide${number}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="${rels}">${[1, 2].map((number) => `<Relationship Id="rId${number}" Type="${rel}/slide" Target="slides/slide${number}.xml"/>`).join('')}</Relationships>`);
  const mediaName = hiddenFromScan ? 'picture.bin' : 'picture.emf';
  for (const number of [1, 2]) {
    const picture = number === pictureSlide ? '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Metafile"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' : '';
    zip.file(`ppt/slides/slide${number}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${picture}</p:spTree></p:cSld></p:sld>`);
    if (picture) zip.file(`ppt/slides/_rels/slide${number}.xml.rels`, `<Relationships xmlns="${rels}"><Relationship Id="picture" Type="${rel}/image" Target="../media/${mediaName}"/></Relationships>`);
  }
  if (pictureSlide) zip.file(`ppt/media/${mediaName}`, await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/metafile-text.emf', import.meta.url)));
  return zip.generateAsync({ type: 'uint8array' });
}

function connection() {
  const pair = createInProcessPair();
  let openId: number | undefined;
  let acknowledged = false;
  const messages: unknown[] = [];
  const host = createPresentationSessionHost({ ...pair.host, post(message: unknown, transfer?: Transferable[]) {
    if (isHostMessage(message) && message.kind === 'reply' && message.id === openId && message.ok) acknowledged = true;
    messages.push(message);
    pair.host.post(message, transfer);
  } }, { initWasm: async () => {} });
  return { host, messages, get acknowledged() { return acknowledged; }, transport: {
    ...pair.client, post(message: unknown, transfer?: Transferable[]) {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'openCommandEditor') openId = message.id;
      pair.client.post(message, transfer);
    },
  } };
}

function loader(load: () => Promise<MetafileSvgModule>) {
  const spy = spyOn(commandMetafileSvgModule, 'load').mockImplementation(load);
  restorers.push(() => spy.mockRestore());
  return spy;
}

const options = { session: 'metafiles', incarnation: 'metafiles-worker', defaultFonts: false };

beforeAll(async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_lean_bg.wasm', import.meta.url)));
});

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.dispose()));
  for (const restore of restorers.splice(0).reverse()) restore();
});

test('a package without metafiles never loads the converter', async () => {
  const load = loader(async () => ({ metafileToSvg: () => svg }));
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(), options, wire.transport);
  clients.push(client);
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(load).not.toHaveBeenCalled();
});

test('opening metafiles are converted before the first reply with default fonts disabled', async () => {
  const pending = deferred<MetafileSvgModule>();
  const requested = deferred<void>();
  loader(() => { requested.resolve(); return pending.promise; });
  const wire = connection();
  const opening = createPresentationCommandClient(await deck(1), options, wire.transport);
  await requested.promise;
  expect(wire.acknowledged).toBe(false);
  pending.resolve({ metafileToSvg: () => svg });
  const client = await opening;
  clients.push(client);
  expect(client.state.frame?.media).toEqual([{ assetId: 'ppt/media/picture.emf', bytes: svg.buffer }]);
});

test('metafiles only on slide two load at idle after the first reply', async () => {
  const requested = deferred<void>();
  const wire = connection();
  loader(async () => {
    expect(wire.acknowledged).toBe(true);
    requested.resolve();
    return { metafileToSvg: () => svg };
  });
  const client = await createPresentationCommandClient(await deck(2), options, wire.transport);
  clients.push(client);
  expect(client.state.frame?.media).toEqual([]);
  await requested.promise;
});

for (const shared of [false, true]) test(`fallback loading sends a worker repaint with fresh SVG media (${shared ? 'shared' : 'direct'})`, async () => {
  const pending = deferred<MetafileSvgModule>();
  const requested = deferred<void>();
  loader(() => { requested.resolve(); return pending.promise; });
  const wire = connection();
  const create = shared ? createSharedPresentationCommandClient : createPresentationCommandClient;
  const client = await create(await deck(1, true), options, wire.transport);
  clients.push(client);
  const first = client.state.frame!;
  expect(first.media[0]?.bytes.byteLength).toBe(0);
  await requested.promise;
  const repainted = deferred<void>();
  client.subscribe((ack) => {
    if (ack.frame && ack.frame.slideRevision > first.slideRevision) repainted.resolve();
  });
  pending.resolve({ metafileToSvg: () => svg });
  await repainted.promise;
  const frame = client.state.frame!;
  expect(frame.slideRevision).toBeGreaterThan(first.slideRevision);
  expect(frame.media[0]?.assetId).toBe(first.media[0]?.assetId);
  expect(frame.media[0]?.bytes).not.toBe(first.media[0]?.bytes);
  expect(new Uint8Array(frame.media[0]!.bytes)).toEqual(svg);
  expect(wire.messages.some((message) => isHostMessage(message) && message.kind === 'event' && message.name === 'defaultFonts')).toBe(true);
});

test('disposal during a converter load sends no later messages', async () => {
  const pending = deferred<MetafileSvgModule>();
  const requested = deferred<void>();
  loader(() => { requested.resolve(); return pending.promise; });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(1, true), options, wire.transport);
  clients.push(client);
  await requested.promise;
  await client.dispose();
  const count = wire.messages.length;
  pending.resolve({ metafileToSvg: () => svg });
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(wire.messages).toHaveLength(count);
});

test('a failed opening load retries after the first frame and repaints', async () => {
  const pending = deferred<MetafileSvgModule>();
  const requested = deferred<void>();
  let attempts = 0;
  loader(() => {
    if (++attempts === 1) return Promise.reject(new Error('unavailable'));
    requested.resolve();
    return pending.promise;
  });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(1), options, wire.transport);
  clients.push(client);
  expect(client.state.frame?.media[0]?.bytes.byteLength).toBe(0);
  await requested.promise;
  const repainted = deferred<void>();
  client.subscribe(() => repainted.resolve());
  pending.resolve({ metafileToSvg: () => svg });
  await repainted.promise;
  expect(new Uint8Array(client.state.frame!.media[0]!.bytes)).toEqual(svg);
});

test('an opening load times out and later repaints without another client request', async () => {
  const pending = deferred<MetafileSvgModule>();
  const load = loader(() => pending.promise);
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(1), options, wire.transport);
  clients.push(client);
  expect(client.state.frame?.media[0]?.bytes.byteLength).toBe(0);
  expect(load).toHaveBeenCalledTimes(1);
  const repainted = deferred<void>();
  client.subscribe(() => repainted.resolve());
  pending.resolve({ metafileToSvg: () => svg });
  await repainted.promise;
  expect(new Uint8Array(client.state.frame!.media[0]!.bytes)).toEqual(svg);
}, 10_000);

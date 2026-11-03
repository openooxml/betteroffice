import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { initWasm } from '../wasm/loader';
import { engineKind } from '../wasm/engine';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from './commandClient';
import { commandFontProvider } from './commandDefaultFonts';
import { openPresentationSession, type PresentationSession } from './client';
import { createPresentationSessionHost } from './host';
import { commandMetafileSvgModule, type MetafileSvgModule } from './metafileSvg';
import type { PresentationCommandClient } from './commandProtocol';

const clients: Array<PresentationCommandClient | PresentationSession> = [];
const restorers: (() => void)[] = [];
const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96"/>');

function svgModule(): MetafileSvgModule {
  return { MetafileConverter: class { toSvg() { return svg; } } };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function deck(pictureSlide?: number, hiddenFromScan = false, slideTwoFont?: string): Promise<Uint8Array> {
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
    const text = number === 2 && slideTwoFont ? `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="${slideTwoFont}"/></a:rPr><a:t>alpha</a:t></a:r></a:p></p:txBody></p:sp>` : '';
    zip.file(`ppt/slides/slide${number}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${picture}${text}</p:spTree></p:cSld></p:sld>`);
    if (picture) zip.file(`ppt/slides/_rels/slide${number}.xml.rels`, `<Relationships xmlns="${rels}"><Relationship Id="picture" Type="${rel}/image" Target="../media/${mediaName}"/></Relationships>`);
  }
  if (pictureSlide) zip.file(`ppt/media/${mediaName}`, await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/metafile-text.emf', import.meta.url)));
  return zip.generateAsync({ type: 'uint8array' });
}

type HostOptions = NonNullable<Parameters<typeof createPresentationSessionHost>[1]>;

function connection(workerRender?: HostOptions['workerRender']) {
  const pair = createInProcessPair();
  let openId: number | undefined;
  let acknowledged = false;
  const messages: unknown[] = [];
  const host = createPresentationSessionHost({ ...pair.host, post(message: unknown, transfer?: Transferable[]) {
    if (isHostMessage(message) && message.kind === 'reply' && message.id === openId && message.ok) acknowledged = true;
    messages.push(message);
    pair.host.post(message, transfer);
  } }, { initWasm: async () => {}, workerRender });
  return { host, messages, get acknowledged() { return acknowledged; }, transport: {
    ...pair.client, post(message: unknown, transfer?: Transferable[]) {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'openCommandEditor') openId = message.id;
      pair.client.post(message, transfer);
    },
  } };
}

function sessionWorker(): Worker {
  const pair = createInProcessPair();
  createPresentationSessionHost(pair.host, { initWasm: async () => {} });
  const target = new EventTarget();
  pair.client.listen((data) => target.dispatchEvent(new MessageEvent('message', { data })));
  return Object.assign(target, {
    postMessage: (message: unknown, transfer?: Transferable[]) => pair.client.post(message, transfer),
    terminate: () => pair.client.close(),
  }) as unknown as Worker;
}

function loader(load: () => Promise<MetafileSvgModule>) {
  const spy = spyOn(commandMetafileSvgModule, 'load').mockImplementation(load);
  restorers.push(() => spy.mockRestore());
  return spy;
}

const options = { session: 'metafiles', incarnation: 'metafiles-worker', defaultFonts: false };
const wasmTest = test.skipIf(engineKind === 'native');
const isSvg = (bytes: ArrayBuffer | Uint8Array | undefined) =>
  bytes !== undefined && /^\s*<svg[\s>]/.test(new TextDecoder().decode(new Uint8Array(bytes).subarray(0, 256)));

beforeAll(async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_lean_bg.wasm', import.meta.url)));
});

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.dispose()));
  for (const restore of restorers.splice(0).reverse()) restore();
});

wasmTest('openPresentationSession frames await lazy SVG media on the lean core', async () => {
  const pending = deferred<MetafileSvgModule>();
  const requested = deferred<void>();
  const load = loader(() => { requested.resolve(); return pending.promise; });
  const client = await openPresentationSession(await deck(1), { worker: sessionWorker });
  clients.push(client);
  const framing = client.call.frame(0);
  await requested.promise;
  pending.resolve(svgModule());
  const first = await framing;
  expect(first.media.get('ppt/media/picture.emf')).toEqual(svg);
  expect((await client.call.frame(0)).media.get('ppt/media/picture.emf')).toEqual(svg);
  expect(load).toHaveBeenCalledTimes(1);
});

wasmTest('openPresentationSession retries a failed converter load on the next frame', async () => {
  let attempts = 0;
  const load = loader(async () => {
    if (++attempts === 1) throw new Error('unavailable');
    return svgModule();
  });
  const client = await openPresentationSession(await deck(1), { worker: sessionWorker });
  clients.push(client);
  expect((await client.call.frame(0)).media.get('ppt/media/picture.emf')).toHaveLength(0);
  expect(load).toHaveBeenCalledTimes(1);
  expect((await client.call.frame(0)).media.get('ppt/media/picture.emf')).toEqual(svg);
  expect(load).toHaveBeenCalledTimes(2);
});

wasmTest('openPresentationSession holds a frame until a slow converter load settles', async () => {
  const pending = deferred<MetafileSvgModule>();
  const load = loader(() => pending.promise);
  const client = await openPresentationSession(await deck(1), { worker: sessionWorker });
  clients.push(client);
  let settled = false;
  const framing = client.call.frame(0).then((frame) => { settled = true; return frame; });
  await new Promise((resolve) => setTimeout(resolve, 3500));
  expect(settled).toBe(false);
  pending.resolve(svgModule());
  expect((await framing).media.get('ppt/media/picture.emf')).toEqual(svg);
  expect(load).toHaveBeenCalledTimes(1);
}, 10_000);

test('a package without metafiles never loads the converter', async () => {
  const load = loader(async () => svgModule());
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(), options, wire.transport);
  clients.push(client);
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(load).not.toHaveBeenCalled();
});

wasmTest('opening metafiles are converted before the first reply with default fonts disabled', async () => {
  const pending = deferred<MetafileSvgModule>();
  const requested = deferred<void>();
  loader(() => { requested.resolve(); return pending.promise; });
  const wire = connection();
  const opening = createPresentationCommandClient(await deck(1), options, wire.transport);
  await requested.promise;
  expect(wire.acknowledged).toBe(false);
  pending.resolve(svgModule());
  const client = await opening;
  clients.push(client);
  expect(client.state.frame?.media).toEqual([{ assetId: 'ppt/media/picture.emf', bytes: svg.buffer }]);
});

wasmTest('metafiles only on slide two load at idle after the first reply', async () => {
  const requested = deferred<void>();
  const wire = connection();
  loader(async () => {
    expect(wire.acknowledged).toBe(true);
    requested.resolve();
    return svgModule();
  });
  const client = await createPresentationCommandClient(await deck(2), options, wire.transport);
  clients.push(client);
  expect(client.state.frame?.media).toEqual([]);
  await requested.promise;
});

for (const shared of [false, true]) wasmTest(`fallback loading sends a worker repaint with fresh SVG media (${shared ? 'shared' : 'direct'})`, async () => {
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
  pending.resolve(svgModule());
  await repainted.promise;
  const frame = client.state.frame!;
  expect(frame.slideRevision).toBeGreaterThan(first.slideRevision);
  expect(frame.media[0]?.assetId).toBe(first.media[0]?.assetId);
  expect(frame.media[0]?.bytes).not.toBe(first.media[0]?.bytes);
  expect(new Uint8Array(frame.media[0]!.bytes)).toEqual(svg);
  expect(wire.messages.some((message) => isHostMessage(message) && message.kind === 'event' && message.name === 'defaultFonts')).toBe(true);
});

wasmTest('shared navigation folds a metafile repaint into its font-gated reply', async () => {
  const converter = deferred<MetafileSvgModule>();
  const converterRequested = deferred<void>();
  const load = loader(() => { converterRequested.resolve(); return converter.promise; });
  const font = deferred<ArrayBuffer>();
  const fontRequested = deferred<void>();
  const provider = spyOn(commandFontProvider, 'load').mockResolvedValue({
    resolve: (family: string) => family === 'Cambria' ? () => { fontRequested.resolve(); return font.promise; } : undefined,
    resolveFamily: () => undefined,
    resolveScriptFallback: () => undefined,
    resolveLastResort: () => { throw new Error('Unrelated substitution'); },
  });
  restorers.push(() => provider.mockRestore());
  const fetch = spyOn(globalThis, 'fetch').mockRejectedValue(new Error('Network disabled'));
  restorers.push(() => fetch.mockRestore());
  const fontBytes = new Uint8Array(await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))).buffer;
  const wire = connection();
  const client = await createSharedPresentationCommandClient(await deck(2, true, 'Cambria'), {
    ...options, defaultFonts: true,
  }, wire.transport);
  clients.push(client);
  await fontRequested.promise;
  expect(load).not.toHaveBeenCalled();
  const destination = client.state.slides[1].id;
  const first = await client.thumbnail(destination);
  expect(first.status).toBe('frame');
  if (first.status !== 'frame') throw new Error('Missing slide two thumbnail');
  expect(first.frame.media[0]?.bytes.byteLength).toBe(0);
  await converterRequested.promise;
  const throughSeq = client.state.throughSeq;
  const acknowledgements: number[] = [];
  let following: ReturnType<PresentationCommandClient['dispatch']> | undefined;
  client.subscribe((ack) => {
    acknowledgements.push(ack.throughSeq);
    if (ack.throughSeq === throughSeq + 1 && ack.outcomes[0]?.status === 'applied') {
      following = client.dispatch({ kind: 'activateSlide', payload: { slideId: client.state.slides[0].id } });
    }
  });
  const start = wire.messages.length;
  let navigated = false;
  const navigation = client.dispatch({ kind: 'activateSlide', payload: { slideId: destination } }).then((ack) => {
    navigated = true;
    return ack;
  });
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(navigated).toBe(false);
  converter.resolve(svgModule());
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(navigated).toBe(false);
  expect(client.state.throughSeq).toBe(throughSeq);
  expect(acknowledgements).toEqual([]);
  expect(wire.messages.slice(start).filter((message) => isHostMessage(message) &&
    (message.kind === 'reply' || message.kind === 'event' && message.name === 'defaultFonts'))).toEqual([]);
  font.resolve(fontBytes);
  const ack = await navigation;
  expect(ack.outcomes[0]?.status).toBe('applied');
  expect(ack.dirtySlideIds).toEqual(client.state.slides.map((slide) => slide.id));
  expect(ack.frame?.slideId).toBe(destination);
  expect(ack.frame?.slideRevision).toBeGreaterThan(first.frame.slideRevision);
  expect(ack.frame?.media).toEqual([{ assetId: 'ppt/media/picture.bin', bytes: svg.buffer }]);
  expect(following).toBeDefined();
  const second = await following!;
  expect(second.outcomes[0]?.status).toBe('applied');
  expect(second.outcomes[0]?.reason).toBeUndefined();
  expect(acknowledgements).toEqual([throughSeq + 1, throughSeq + 2]);
  expect(wire.messages.slice(start).some((message) => isHostMessage(message) &&
    message.kind === 'event' && message.name === 'defaultFonts')).toBe(false);
  expect(load).toHaveBeenCalledTimes(1);
});

wasmTest('disposal during a converter load sends no later messages', async () => {
  const pending = deferred<MetafileSvgModule>();
  const requested = deferred<void>();
  loader(() => { requested.resolve(); return pending.promise; });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(1, true), options, wire.transport);
  clients.push(client);
  await requested.promise;
  await client.dispose();
  const count = wire.messages.length;
  pending.resolve(svgModule());
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(wire.messages).toHaveLength(count);
});

wasmTest('a failed opening load retries after the first frame and repaints', async () => {
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
  pending.resolve(svgModule());
  await repainted.promise;
  expect(new Uint8Array(client.state.frame!.media[0]!.bytes)).toEqual(svg);
});

wasmTest('a late metafile repaint drops the worker raster media cached from the empty first frame', async () => {
  const pending = deferred<MetafileSvgModule>();
  loader(() => pending.promise);
  const events: string[] = [];
  const frames: { mediaBytes(assetId: string): Uint8Array }[] = [];
  const workerRender = {
    transport: undefined as never, active: () => true, pending: () => false, afterReply() {},
    paintNow: async () => undefined, defaultFonts() {}, dispose() {},
    frame(frame: { mediaBytes(assetId: string): Uint8Array }) { frames.push(frame); events.push('frame'); },
    resetMedia() { events.push('reset'); },
  } as unknown as NonNullable<Parameters<typeof connection>[0]>;
  const wire = connection(workerRender);
  const client = await createPresentationCommandClient(await deck(1), options, wire.transport);
  clients.push(client);
  expect(frames[0]?.mediaBytes('ppt/media/picture.emf').byteLength).toBe(0);
  expect(events).not.toContain('reset');
  const repainted = deferred<void>();
  client.subscribe(() => repainted.resolve());
  pending.resolve(svgModule());
  await repainted.promise;
  const reset = events.indexOf('reset');
  expect(reset).toBeGreaterThan(-1);
  expect(events.lastIndexOf('frame')).toBeGreaterThan(reset);
  expect(isSvg(frames[frames.length - 1]!.mediaBytes('ppt/media/picture.emf'))).toBe(true);
}, 10_000);

wasmTest('the worker raster takes each extracted picture buffer once and the frame keeps no copy', async () => {
  loader(async () => svgModule());
  const zip = await JSZip.loadAsync(await deck(1));
  const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));
  zip.file('ppt/media/picture.emf', png);
  const frames: { mediaBytes(assetId: string): Uint8Array }[] = [];
  const workerRender = {
    transport: undefined as never, active: () => true, pending: () => false, afterReply() {},
    paintNow: async () => undefined, defaultFonts() {}, dispose() {}, resetMedia() {},
    frame(frame: { mediaBytes(assetId: string): Uint8Array }) { frames.push(frame); },
  } as unknown as NonNullable<Parameters<typeof connection>[0]>;
  const wire = connection(workerRender);
  const client = await createPresentationCommandClient(await zip.generateAsync({ type: 'uint8array' }), options, wire.transport);
  clients.push(client);
  const first = frames[0]!.mediaBytes('ppt/media/picture.emf');
  const again = frames[0]!.mediaBytes('ppt/media/picture.emf');
  expect(first).toEqual(png);
  expect(again).toEqual(png);
  expect(again).not.toBe(first);
});

wasmTest('an opening load times out and later repaints without another client request', async () => {
  const pending = deferred<MetafileSvgModule>();
  const load = loader(() => pending.promise);
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(1), options, wire.transport);
  clients.push(client);
  expect(client.state.frame?.media[0]?.bytes.byteLength).toBe(0);
  expect(load).toHaveBeenCalledTimes(1);
  const repainted = deferred<void>();
  client.subscribe(() => repainted.resolve());
  pending.resolve(svgModule());
  await repainted.promise;
  expect(new Uint8Array(client.state.frame!.media[0]!.bytes)).toEqual(svg);
}, 10_000);

test('an opening-slide metafile reaches the first editor frame as SVG on every engine', async () => {
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(1), options, wire.transport);
  clients.push(client);
  const picture = client.state.frame?.media.find(({ assetId }) => assetId === 'ppt/media/picture.emf');
  expect(isSvg(picture?.bytes)).toBe(true);
});

test('a public session frame carries a metafile picture as SVG on every engine', async () => {
  const client = await openPresentationSession(await deck(1), { worker: sessionWorker });
  clients.push(client);
  expect(isSvg((await client.call.frame(0)).media.get('ppt/media/picture.emf'))).toBe(true);
});

import { afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { defaultFontInstaller, type DefaultFontFace, type DefaultFontPayload } from '../../../../shared/pptxDefaultFonts';
import { initWasm } from '../wasm/loader';
import type { SlideDisplayList } from '../types';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from './commandClient';
import { commandFontProvider } from './commandDefaultFonts';
import { createPresentationSessionHost } from './host';
import type { PresentationCommandClient, PresentationCommandOpenOptions } from './commandProtocol';

const restorers: (() => void)[] = [];
const clients: PresentationCommandClient[] = [];
let font: ArrayBuffer;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function deck(families = ['Calibri', 'Cambria'], text = 'alpha'): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${families.map((_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${families.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${families.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${rel}/slide" Target="slides/slide${index + 1}.xml"/>`).join('')}</Relationships>`);
  for (const [index, family] of families.entries()) zip.file(`ppt/slides/slide${index + 1}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="${family}"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function source(loaders: Record<string, () => Promise<ArrayBuffer>>) {
  const resolve = mock((family: string) => loaders[family]);
  const imported = spyOn(commandFontProvider, 'load').mockResolvedValue({ resolve, resolveFamily: () => undefined,
    resolveScriptFallback: () => undefined, resolveLastResort: () => { throw new Error('Unrelated substitution'); } });
  restorers.push(() => imported.mockRestore());
  const fetch = spyOn(globalThis, 'fetch').mockRejectedValue(new Error('Network disabled'));
  restorers.push(() => fetch.mockRestore());
  return { imported, resolve, fetch };
}

function connection() {
  const pair = createInProcessPair();
  const messages: unknown[] = [];
  let openId: number | undefined;
  let acknowledged = false;
  const requested = deferred<void>();
  const host = { ...pair.host, post(message: unknown, transfer?: Transferable[]) {
    if (isHostMessage(message) && message.kind === 'reply' && message.id === openId && message.ok) acknowledged = true;
    messages.push(message);
    pair.host.post(message, transfer);
  } };
  createPresentationSessionHost(host, { initWasm: async () => {} });
  return { messages, requested: requested.promise, get acknowledged() { return acknowledged; }, transport: {
    ...pair.client, post(message: unknown, transfer?: Transferable[]) {
      if (isClientMessage(message) && message.kind === 'call') {
        if (message.method === 'openCommandEditor') openId = message.id;
        if (message.method === 'command' || message.method === 'edit') requested.resolve();
      }
      pair.client.post(message, transfer);
    },
  } };
}

const options = { session: 'default-fonts', incarnation: 'worker-fonts', fonts: [] } satisfies PresentationCommandOpenOptions;

beforeAll(async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  font = new Uint8Array(await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))).buffer;
});

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.dispose()));
  for (const restore of restorers.splice(0).reverse()) restore();
});

test('the initial ack waits for active fonts and main installation; background fonts arrive after ack', async () => {
  const active = deferred<ArrayBuffer>();
  const background = deferred<ArrayBuffer>();
  const activeRequested = deferred<void>();
  const backgroundRequested = deferred<void>();
  source({ Calibri: () => { activeRequested.resolve(); return active.promise; },
    Cambria: () => { backgroundRequested.resolve(); return background.promise; } });
  const wire = connection();
  const installed = deferred<void>();
  const installing = deferred<void>();
  const faces: DefaultFontFace[] = [];
  const opening = createPresentationCommandClient(await deck(), Object.assign({ ...options, readOnly: true }, {
    [defaultFontInstaller]: async (additions: readonly DefaultFontFace[]) => {
      faces.push(...additions);
      installing.resolve();
      await installed.promise;
    },
  }), wire.transport);
  await activeRequested.promise;
  expect(wire.acknowledged).toBe(false);
  active.resolve(font);
  await installing.promise;
  let ready = false;
  void opening.then(() => { ready = true; });
  expect(ready).toBe(false);
  expect(faces.map((face) => face.family)).toEqual(['Calibri']);
  installed.resolve();
  const client = await opening;
  clients.push(client);
  expect(client.state.frame?.slideIndex).toBe(0);
  await backgroundRequested.promise;
  const arrived = deferred<void>();
  client.subscribe((ack) => { if ((ack as DefaultFontPayload).defaultFonts?.some((face) => face.family === 'Cambria')) arrived.resolve(); });
  background.resolve(font);
  await arrived.promise;
  expect(faces.map((face) => face.family)).toEqual(['Calibri', 'Cambria']);
  expect(wire.messages.some((message) => isHostMessage(message) && message.kind === 'event' && message.name === 'defaultFonts')).toBe(true);
  expect(font.byteLength).toBeGreaterThan(0);
});

for (const shared of [false, true]) test(`navigation waits only for target pending fonts (${shared ? 'shared' : 'direct'})`, async () => {
  const target = deferred<ArrayBuffer>();
  const targetRequested = deferred<void>();
  source({ Calibri: async () => font, Cambria: () => { targetRequested.resolve(); return target.promise; },
    Arial: () => new Promise(() => {}) });
  const wire = connection();
  const create = shared ? createSharedPresentationCommandClient : createPresentationCommandClient;
  const client = await create(await deck(['Calibri', 'Cambria', 'Arial']), options, wire.transport);
  clients.push(client);
  await targetRequested.promise;
  const destination = client.state.slides[1].id;
  let navigated = false;
  const navigation = client.dispatch({ kind: 'activateSlide', payload: { slideId: destination } }).then((ack) => {
    navigated = true;
    return ack;
  });
  await wire.requested;
  expect(navigated).toBe(false);
  expect(client.state.activeSlideId).not.toBe(destination);
  target.resolve(font);
  const ack = await navigation;
  expect(ack.frame?.slideId).toBe(destination);
  expect(ack.outcomes[0].status).toBe('applied');
});

test('a supplied family excludes defaults and all-covered decks never import the manager', async () => {
  const manager = source({ Calibri: async () => font });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Calibri']), {
    ...options, fonts: [{ family: ' calibri ', bold: true, bytes: font.slice(0) }],
  }, wire.transport);
  clients.push(client);
  expect(manager.imported).not.toHaveBeenCalled();
  expect(manager.fetch).not.toHaveBeenCalled();
});

test('failed and stalled defaults still acknowledge an open', async () => {
  source({ Calibri: async () => { throw new Error('Missing asset'); }, Cambria: () => new Promise(() => {}) });
  const timeout = globalThis.setTimeout;
  const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: TimerHandler, delay?: number, ...args: unknown[]) =>
    timeout(callback, delay === 5000 ? 10 : delay, ...args)) as typeof setTimeout);
  restorers.push(() => timer.mockRestore());
  for (const family of ['Calibri', 'Cambria']) {
    const wire = connection();
    const client = await createPresentationCommandClient(await deck([family]), options, wire.transport);
    clients.push(client);
    expect(wire.acknowledged).toBe(true);
    expect((client.state as DefaultFontPayload).defaultFonts).toEqual([]);
    expect(client.state.frame).not.toBeNull();
  }
});

test('formatting updates collect new families in the background', async () => {
  const requested = deferred<void>();
  source({ Calibri: async () => font, Cambria: async () => { requested.resolve(); return font; } });
  const wire = connection();
  const client = await createSharedPresentationCommandClient(await deck(['Calibri']), options, wire.transport);
  clients.push(client);
  const content = (await client.read()).result;
  if (!content.ok) throw new Error(content.failure.message);
  const story = content.stories[0];
  const arrived = deferred<void>();
  client.subscribe((ack) => { if ((ack as DefaultFontPayload).defaultFonts?.some((face) => face.family === 'Cambria')) arrived.resolve(); });
  await client.dispatch({ kind: 'formatText', payload: { inputTime: 1, undoGroup: 'fonts', patch: { fontFamily: 'Cambria' },
    at: { source: 'confirmed', selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
      anchor: 0, focus: 5, affinity: 'downstream' } } } });
  await requested.promise;
  await arrived.promise;
});

test('uncovered Devanagari keeps browser fallback after a default face registers', async () => {
  const carlito = new Uint8Array(await readFile(new URL('../../../fonts/assets/Carlito-Regular.ttf', import.meta.url))).buffer;
  source({ Calibri: async () => carlito });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Calibri'], 'alpha नमस्ते'), options, wire.transport);
  clients.push(client);
  const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
  const boxes = list.primitives.filter((primitive) => primitive.kind === 'textBox');
  const runs = boxes.flatMap((box) => box.lines.flatMap((line) => line.runs));
  expect(runs.some((run) => run.text.includes('alpha') && run.fontId !== 0xffffffff)).toBe(true);
  expect(runs.filter((run) => /[\u0900-\u097f]/.test(run.text)).length).toBeGreaterThan(0);
  expect(runs.filter((run) => /[\u0900-\u097f]/.test(run.text)).every((run) => run.fontId === 0xffffffff)).toBe(true);
});

test('registered coverage fonts still precede the browser fallback', async () => {
  const arabic = new Uint8Array(await readFile(new URL('../../../fonts/assets/NotoSansArabic-Regular.ttf', import.meta.url))).buffer;
  source({ Calibri: async () => font });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Calibri'], 'alpha مرحبا नमस्ते'), {
    ...options, fallbackFonts: [{ family: 'Noto Sans Arabic', bytes: arabic }],
  }, wire.transport);
  clients.push(client);
  const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
  const boxes = list.primitives.filter((primitive) => primitive.kind === 'textBox');
  const runs = boxes.flatMap((box) => box.lines.flatMap((line) => line.runs));
  expect(runs.some((run) => /[\u0600-\u06ff]/.test(run.text) && run.fontFamily === 'Noto Sans Arabic' && run.fontId !== 0xffffffff)).toBe(true);
  expect(runs.filter((run) => /[\u0900-\u097f]/.test(run.text)).every((run) => run.fontId === 0xffffffff)).toBe(true);
});

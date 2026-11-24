import { afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import * as fonts from '@betteroffice/fonts';
import * as progressive from '@betteroffice/fonts/progressive';
import { fontAssetResponse, withBrowserFontGlobals } from '../../../fonts/src/fontAssets.test-support';
import { spyOnFontFetch } from '../../../fonts/src/fontFetch.test-support';
import { createFontRuntime } from '../../../fonts/src/fontRuntime';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import type { FontCdnDiagnostic } from '../../../../shared/fontCdnDiagnostics';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { engineKind } from '../wasm/enginePng';
import { defaultFontInstaller, type DefaultFontFace, type DefaultFontPayload } from '../../../../shared/pptxDefaultFonts';
import { decodeFontFace, initWasm } from '../wasm/loader';
import { PptxRenderer } from '../wasm/engine';
import type { SlideDisplayList } from '../types';
import { openPresentationSession, type PresentationSession } from '../worker';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from './commandClient';
import { commandFontProvider } from './commandDefaultFonts';
import { createPresentationSessionHost } from './host';
import type { PresentationCommandClient, PresentationCommandOpenOptions } from './commandProtocol';
import * as sessionWorker from './sessionWorker';

// Match main's measured WASM PNG behavior with the optional pack absent.
const EXPORT_WITH_ABSENT_PACK: 'draws' | 'font-unavailable' = 'draws';

async function notoSansJP(): Promise<ArrayBuffer> {
  const packed = await readFile(new URL('../../../fonts-cjk/assets/NotoSansJP-Regular.otf.zst', import.meta.url));
  return (await decodeFontFace(packed, { codec: 'zstd' })).font;
}

const restorers: (() => void)[] = [];
const clients: Array<PresentationCommandClient | PresentationSession> = [];
let font: ArrayBuffer;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function deck(families = ['Calibri', 'Cambria'], text = 'alpha', traits = ''): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${families.map((_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${families.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${families.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${rel}/slide" Target="slides/slide${index + 1}.xml"/>`).join('')}</Relationships>`);
  for (const [index, family] of families.entries()) zip.file(`ppt/slides/slide${index + 1}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" ${traits}><a:latin typeface="${family}"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function source(
  loaders: Record<string, () => Promise<ArrayBuffer>>,
  script: (script: string, bold: boolean, italic: boolean) => (() => Promise<ArrayBuffer>) | undefined = () => undefined,
) {
  const resolve = mock((family: string, _bold?: boolean, _italic?: boolean) => loaders[family]);
  const resolveScriptFallback = mock(script);
  const imported = spyOn(commandFontProvider, 'load').mockResolvedValue({ resolve, resolveFamily: () => undefined,
    resolveScriptFallback, resolveLastResort: () => { throw new Error('Unrelated substitution'); } });
  restorers.push(() => imported.mockRestore());
  const fetch = spyOn(globalThis, 'fetch').mockRejectedValue(new Error('Network disabled'));
  restorers.push(() => fetch.mockRestore());
  return { imported, resolve, resolveScriptFallback, fetch };
}

function connection() {
  const pair = createInProcessPair();
  const messages: unknown[] = [];
  const posts: Array<{ message: unknown; transfer: Transferable[] }> = [];
  let openId: number | undefined;
  let acknowledged = false;
  const requested = deferred<void>();
  const host = { ...pair.host, post(message: unknown, transfer?: Transferable[]) {
    if (isHostMessage(message) && message.kind === 'reply' && message.id === openId && message.ok) acknowledged = true;
    messages.push(message);
    posts.push({ message, transfer: transfer ?? [] });
    pair.host.post(message, transfer);
  } };
  createPresentationSessionHost(host, { initWasm: async () => {} });
  return { messages, posts, requested: requested.promise, get acknowledged() { return acknowledged; }, transport: {
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

async function scriptCalls(manager: ReturnType<typeof source>, count: number): Promise<string[]> {
  for (let tries = 0; tries < 400 && manager.resolveScriptFallback.mock.calls.length < count; tries++) {
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  return manager.resolveScriptFallback.mock.calls.map((call) => JSON.stringify(call)).sort();
}

const expectedCalls = (...list: unknown[][]) => list.map((call) => JSON.stringify(call)).sort();

beforeAll(async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  font = new Uint8Array(await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))).buffer;
});

afterEach(async () => {
  await Promise.all(clients.splice(0).map((client) => client.dispose()));
  for (const restore of restorers.splice(0).reverse()) restore();
});

function stubPngCanvas(encoded: BlobPart = 'png') {
  const painted = mock((_text: string) => {});
  class Canvas {
    getContext() {
      return new Proxy({ fillText: painted, measureText: (text: string) => ({ width: text.length }) }, {
        get: (context, property) => Reflect.get(context, property) ?? (() => {}),
      });
    }
    async convertToBlob() { return new Blob([encoded], { type: 'image/png' }); }
  }
  class Face { async load() { return this; } }
  const globals = { document: undefined, OffscreenCanvas: Canvas, FontFace: Face, fonts: new Set(),
    navigator: { userAgent: 'Chrome/149.0' } };
  for (const [key, value] of Object.entries(globals)) {
    const previous = Object.getOwnPropertyDescriptor(globalThis, key);
    Object.defineProperty(globalThis, key, { configurable: true, value });
    restorers.push(() => {
      if (previous) Object.defineProperty(globalThis, key, previous);
      else Reflect.deleteProperty(globalThis, key);
    });
  }
  return painted;
}

function blockStandaloneFontCdn(latinInstalled: boolean) {
  const previous = getFontCdnConfiguration();
  const restoreBrowser = withBrowserFontGlobals();
  setFontCdnConfiguration(undefined);
  const runtime = createFontRuntime(async (pack) => ({
    cdn: getFontCdnConfiguration(),
    urls: pack === 'base' || (pack === 'latin' && latinInstalled)
      ? Object.fromEntries(Object.entries((await fonts.fontAssetUrls(pack))!).map(([file, href]) =>
        [file, () => new URL(href)])) : undefined,
  }));
  const installed = spyOn(fonts, 'bundledFaceInstalled').mockImplementation(runtime.bundledFaceInstalled);
  const load = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(runtime.loadBundledFontBytes);
  const worker = spyOn(sessionWorker, 'createSessionWorker').mockImplementation(() => {
    const pair = createInProcessPair();
    createPresentationSessionHost(pair.host, { initWasm: async () => {} });
    const target = new EventTarget();
    pair.client.listen((data) => target.dispatchEvent(new MessageEvent('message', { data })));
    return Object.assign(target, { postMessage: pair.client.post, terminate: pair.client.close }) as unknown as Worker;
  });
  const fetch = spyOnFontFetch(async (input) => {
    if (new URL(String(input)).hostname === 'cdn.jsdelivr.net') throw new TypeError('Network disabled');
    return fontAssetResponse(input);
  });
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => {
    warning.mockRestore();
    fetch.mockRestore();
    worker.mockRestore();
    load.mockRestore();
    installed.mockRestore();
    restoreBrowser();
    setFontCdnConfiguration(previous);
  });
  return { fetch, warning };
}

const standaloneWasmTest = test.skipIf(engineKind === 'native');

standaloneWasmTest('standalone renderPng matches the absent CJK pack when the default CDN is blocked', async () => {
  const { fetch, warning } = blockStandaloneFontCdn(true);
  const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jxgQAAAABJRU5ErkJggg=='), (byte) => byte.charCodeAt(0));
  const painted = stubPngCanvas(png);
  const register = spyOn(PptxRenderer.prototype, 'registerFont');
  restorers.push(() => register.mockRestore());
  const diagnostics: FontCdnDiagnostic[] = [];
  const dispatch = Object.getOwnPropertyDescriptor(globalThis, 'dispatchEvent');
  Object.defineProperty(globalThis, 'dispatchEvent', {
    configurable: true,
    value(event: CustomEvent<FontCdnDiagnostic>) {
      if (event.type === 'betteroffice:font-pack-diagnostic') diagnostics.push(event.detail);
      return true;
    },
  });
  restorers.push(() => {
    if (dispatch) Object.defineProperty(globalThis, 'dispatchEvent', dispatch);
    else Reflect.deleteProperty(globalThis, 'dispatchEvent');
  });
  const bytes = await deck(['MS Gothic'], 'q漢字');
  fonts.configureFontCdn({ baseUrl: false });
  const absent = await openPresentationSession(bytes);
  clients.push(absent);
  const expectedPng = await absent.call.renderPng({ slide: 0 });
  const primary = () => register.mock.calls.filter(([family]) => family === 'MS Gothic').map(([, , , data]) => ({
    size: data.byteLength, hash: createHash('sha256').update(data).digest('hex'),
  }));
  const expectedPrimary = primary();
  const expectedText = painted.mock.calls.map(([text]) => text).join('');
  register.mockClear();
  painted.mockClear();

  setFontCdnConfiguration(undefined);
  const session = await openPresentationSession(bytes);
  clients.push(session);
  const { call } = session;
  await expect(call.renderPng({ slide: 0 })).resolves.toEqual(expectedPng);
  expect(expectedPng.byteLength).toBeGreaterThan(8);
  expect(expectedPng.byteLength).toBeLessThan(64 * 1024);
  expect([...expectedPng.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  const carlito = fonts.resolveLastResortChoice('MS Gothic', false, false, 'powerpoint').face!;
  expect(expectedPrimary.map(({ size }) => size)).toEqual([carlito.byteLength]);
  expect(primary()).toEqual(expectedPrimary);
  expect(expectedText).toContain('q');
  expect(expectedText).toContain('漢字');
  expect(painted.mock.calls.map(([text]) => text).join('')).toBe(expectedText);

  const diagnostic = expect.objectContaining({ pack: 'cjk', reason: 'cdn-unavailable' });
  expect(diagnostics).toEqual([diagnostic]);
  await expect(call.renderPng({ slide: 0 })).resolves.toEqual(expectedPng);
  expect(diagnostics).toEqual([diagnostic, diagnostic]);
  expect(fetch.mock.calls.map(([url]) => String(url)).filter((url) => url.includes('cdn.jsdelivr.net'))).toEqual([
    expect.stringContaining('/fonts-cjk@'),
  ]);
  expect(warning.mock.calls.filter(([message]) => String(message).includes('Cannot load @betteroffice/fonts-cjk'))).toHaveLength(1);
});

standaloneWasmTest('standalone renderPng rejects when an explicit CJK CDN fails', async () => {
  blockStandaloneFontCdn(true);
  stubPngCanvas();
  fonts.configureFontCdn({ baseUrl: 'https://cdn.jsdelivr.net/npm/' });
  const bytes = await deck(['MS Gothic'], 'q漢字');
  const session = await openPresentationSession(bytes);
  clients.push(session);
  await expect(session.call.renderPng({ slide: 0 })).rejects.toMatchObject({
    code: 'font-unavailable', message: expect.stringContaining('Noto Sans JP for MS Gothic'),
  });
});

standaloneWasmTest('standalone renderPng rejects when an implicit Latin CDN fails', async () => {
  blockStandaloneFontCdn(false);
  stubPngCanvas();
  const bytes = await deck(['Calibri'], 'q');
  const session = await openPresentationSession(bytes);
  clients.push(session);
  await expect(session.call.renderPng({ slide: 0 })).rejects.toMatchObject({
    code: 'font-unavailable', message: expect.stringContaining('Cannot load @betteroffice/fonts-latin'),
  });
});

async function expectPng(client: PresentationCommandClient, outcome: 'draws' | 'font-unavailable', painted: ReturnType<typeof stubPngCanvas>) {
  const png = client.renderPng!({ slide: 0 });
  if (outcome === 'draws' && engineKind === 'native') {
    const bytes = await png;
    expect(bytes.byteLength).toBeGreaterThan(8);
    expect([...bytes.subarray(0, 4)]).toEqual([137, 80, 78, 71]);
  } else if (outcome === 'draws') {
    expect([...await png]).toEqual([...new TextEncoder().encode('png')]);
    expect(painted.mock.calls.map(([text]) => text).join('')).toContain('Xalpha');
  } else await expect(png).rejects.toMatchObject({ code: outcome });
}

for (const [pack, family] of [['cjk', 'MS Gothic'], ['intl', 'Noto Sans Devanagari']] as const) {
  for (const slicedFonts of [false, true]) {
    test(`a blocked default ${pack} CDN matches an absent pack with sliced fonts ${slicedFonts}`, async () => {
      const previous = getFontCdnConfiguration();
      const restoreBrowser = withBrowserFontGlobals();
      restorers.push(() => { restoreBrowser(); setFontCdnConfiguration(previous); });
      setFontCdnConfiguration(undefined);
      const cdn = getFontCdnConfiguration();
      const runtime = createFontRuntime(async (kind) => ({
        cdn: getFontCdnConfiguration(),
        urls: kind === 'base' || kind === 'latin'
          ? Object.fromEntries(Object.entries(await fonts.fontAssetUrls(kind) ?? {}).map(([file, href]) =>
            [file, () => new URL(href)])) : undefined,
      }));
      const provider = spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
      const installed = spyOn(fonts, 'bundledFaceInstalled').mockImplementation(runtime.bundledFaceInstalled);
      const load = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(runtime.loadBundledFontBytes);
      const warn = spyOn(console, 'warn').mockImplementation(() => {});
      const fetched: string[] = [];
      const fetch = spyOnFontFetch(async (input) => {
        const url = new URL(String(input));
        if (url.hostname === 'cdn.jsdelivr.net') {
          fetched.push(url.href);
          throw new TypeError('Network disabled');
        }
        return fontAssetResponse(input);
      });
      const register = spyOn(PptxRenderer.prototype, 'registerFont');
      const sliced = spyOn(progressive, 'openSlicedFace');
      restorers.push(() => { provider.mockRestore(); installed.mockRestore(); load.mockRestore(); warn.mockRestore();
        fetch.mockRestore(); register.mockRestore(); sliced.mockRestore(); });
      const painted = stubPngCanvas();
      const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
      const substitute = fonts.resolveMetricCompatFace('Arial', false, false)!;
      const expected = digest(new Uint8Array(await runtime.loadBundledFontBytes(substitute, undefined, { decoder: decodeFontFace })));
      const wire = connection();
      const client = await createPresentationCommandClient(await deck([family]), { ...options, fontCdn: cdn, slicedFonts }, wire.transport);
      clients.push(client);
      const registrations = () => register.mock.calls.filter(([name]) => name === family)
        .map(([name, bold, italic, bytes]) => [name, bold, italic, digest(bytes)]);
      expect(fetched).toHaveLength(1);
      expect(fetched[0]).toContain(`/fonts-${pack}@`);
      expect(sliced).toHaveBeenCalledTimes(pack === 'cjk' && slicedFonts ? 1 : 0);
      expect(registrations()).toEqual([[family, false, false, expected]]);

      const { result } = await client.read();
      if (!result.ok) throw new Error(result.failure.message);
      const { slideId, shapeId, storyId } = result.stories[0];
      const ack = await client.dispatch({ kind: 'insertText', payload: { inputTime: 1, undoGroup: 'fonts', text: 'X',
        style: { fontFamily: family, bold: false, italic: false },
        at: { source: 'confirmed', selection: { slideId, shapeId, storyId, anchor: 0, focus: 0, affinity: 'downstream' } } } });
      expect(ack.outcomes[0].status).toBe('applied');
      expect(registrations()).toEqual([[family, false, false, expected]]);
      await expectPng(client, EXPORT_WITH_ABSENT_PACK, painted);
      expect(warn.mock.calls.filter(([message]) => String(message).includes(`Cannot load @betteroffice/fonts-${pack}`))).toHaveLength(1);
    });
  }
}

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
  const landings = wire.posts.filter(({ message }) => isHostMessage(message) && (message.kind === 'event'
    ? message.name === 'defaultFonts' : message.kind === 'reply' && message.ok &&
      (message.value as DefaultFontPayload | null)?.defaultFonts !== undefined));
  const landed = landings.flatMap(({ message }) => ((isHostMessage(message) && message.kind === 'event'
    ? message.payload : isHostMessage(message) && message.kind === 'reply' && message.ok ? message.value : null) as
    DefaultFontPayload | null)?.defaultFonts ?? []);
  expect(landed.map((face) => face.family)).toEqual(['Calibri', 'Cambria']);
  for (const face of landed) expect('bytes' in face).toBe(false);
  for (const { transfer } of landings) {
    expect(transfer.some((buffer) => buffer instanceof ArrayBuffer && buffer.byteLength === font.byteLength)).toBe(false);
  }
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

test('app supplied fonts request missing script coverage without replacing primary faces', async () => {
  const manager = source({ Calibri: async () => font });
  const coverage = spyOn(PptxRenderer.prototype, 'setBrowserCoverageFallback');
  restorers.push(() => coverage.mockRestore());
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Calibri'], 'alpha 👩‍💻 こんにちは \uE000 ✓'), {
    ...options, fonts: [{ family: ' calibri ', bold: true, bytes: font.slice(0) }],
  }, wire.transport);
  clients.push(client);
  expect(manager.imported).toHaveBeenCalledTimes(1);
  expect(manager.resolve).not.toHaveBeenCalled();
  expect(manager.resolveScriptFallback).toHaveBeenCalledWith('cjk-jp', false, false);
  expect(manager.fetch).not.toHaveBeenCalled();
  expect(coverage).not.toHaveBeenCalled();
  const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
  const boxes = list.primitives.filter((primitive) => primitive.kind === 'textBox');
  const runs = boxes.flatMap((box) => box.lines.flatMap((line) => line.runs));
  expect(runs.some((run) => run.text.includes('👩‍💻'))).toBe(true);
  expect(runs.every((run) => run.fontId !== 0xffffffff)).toBe(true);
});

test('an app supplying only Arial still gets defaults for the other document families', async () => {
  const manager = source({ Arial: async () => font, 'Times New Roman': async () => font, Calibri: async () => font,
    'MS Mincho': async () => font });
  const register = spyOn(PptxRenderer.prototype, 'registerFont');
  restorers.push(() => register.mockRestore());
  const wire = connection();
  const arial = [false, true].flatMap((bold) => [false, true].map((italic) => ({ family: 'Arial', bold, italic, bytes: font.slice(0) })));
  const faces: string[] = [];
  const client = await createPresentationCommandClient(await deck(['Times New Roman', 'Calibri', 'MS Mincho'], 'alpha \u65e5\u672c'),
    Object.assign({ ...options, fonts: arial }, {
      [defaultFontInstaller]: async (additions: readonly DefaultFontFace[]) => { faces.push(...additions.map((face) => face.family)); },
    }), wire.transport);
  clients.push(client);
  for (const slide of client.state.slides.slice(1)) await client.dispatch({ kind: 'activateSlide', payload: { slideId: slide.id } });
  await client.flush();
  const requested = manager.resolve.mock.calls.map(([family]) => family);
  expect(requested).not.toContain('Arial');
  for (const family of ['Times New Roman', 'Calibri', 'MS Mincho']) {
    expect(requested).toContain(family);
    expect(faces).toContain(family);
  }
  const registered = register.mock.calls.map(([family]) => family);
  expect(registered.filter((family) => family === 'Arial')).toHaveLength(4);
  expect(registered.filter((family) => family !== 'Arial').sort()).toEqual(['Calibri', 'MS Mincho', 'Times New Roman']);
});

test('opting out keeps app supplied sessions free of default fonts', async () => {
  const manager = source({ Calibri: async () => font, Cambria: async () => font });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(), {
    ...options, fonts: [{ family: 'Arial', bytes: font.slice(0) }], defaultFonts: false,
  }, wire.transport);
  clients.push(client);
  await client.dispatch({ kind: 'activateSlide', payload: { slideId: client.state.slides[1].id } });
  await client.flush();
  expect(manager.imported).not.toHaveBeenCalled();
  expect(manager.fetch).not.toHaveBeenCalled();
});

test('default sessions enable browser coverage once before registering faces', async () => {
  source({ Calibri: async () => font, Cambria: async () => font });
  const calls: string[] = [];
  const enable = PptxRenderer.prototype.setBrowserCoverageFallback;
  const coverage = spyOn(PptxRenderer.prototype, 'setBrowserCoverageFallback').mockImplementation(function (
    this: PptxRenderer, enabled: boolean,
  ) {
    calls.push(`coverage:${enabled}`);
    enable.call(this, enabled);
  });
  const original = PptxRenderer.prototype.registerFont;
  const register = spyOn(PptxRenderer.prototype, 'registerFont').mockImplementation(function (
    this: PptxRenderer, family: string, bold: boolean, italic: boolean, bytes: Uint8Array,
  ) {
    calls.push(family);
    return original.call(this, family, bold, italic, bytes);
  });
  restorers.push(() => { register.mockRestore(); coverage.mockRestore(); });
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(), options, wire.transport);
  clients.push(client);
  await client.dispatch({ kind: 'activateSlide', payload: { slideId: client.state.slides[1].id } });
  expect(register.mock.calls.map(([family]) => family)).toEqual(['Calibri', 'Cambria']);
  expect(coverage).toHaveBeenCalledTimes(1);
  expect(calls).toEqual(['coverage:true', 'Calibri', 'Cambria']);
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

test('a stalled coverage scan never holds back the default fonts', async () => {
  const manager = source({ Calibri: async () => font });
  const native = globalThis.DecompressionStream;
  globalThis.DecompressionStream = class {
    readonly readable = new ReadableStream<Uint8Array>();
    readonly writable = new WritableStream<BufferSource>();
  } as unknown as typeof DecompressionStream;
  restorers.push(() => { globalThis.DecompressionStream = native; });
  const timeout = globalThis.setTimeout;
  const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: TimerHandler, delay?: number, ...args: unknown[]) =>
    timeout(callback, delay === 5000 ? 50 : delay, ...args)) as typeof setTimeout);
  restorers.push(() => timer.mockRestore());
  const deflated = await (await JSZip.loadAsync(await deck(['Calibri'])))
    .generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const faces: string[] = [];
  const client = await createPresentationCommandClient(deflated, Object.assign({ ...options }, {
    [defaultFontInstaller]: async (additions: readonly DefaultFontFace[]) => { faces.push(...additions.map((face) => face.family)); },
  }), connection().transport);
  clients.push(client);
  expect(manager.resolve.mock.calls.map(([family]) => family)).toContain('Calibri');
  expect(faces).toEqual(['Calibri']);
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

for (const supplied of [false, true]) test(`script defaults install in the coverage chain with supplied primary ${supplied}`, async () => {
  const japanese = await notoSansJP();
  const manager = source({}, (script) => script === 'cjk-jp' ? async () => japanese : undefined);
  const fallback = spyOn(PptxRenderer.prototype, 'registerFallbackFont');
  const primary = spyOn(PptxRenderer.prototype, 'registerFont');
  restorers.push(() => { fallback.mockRestore(); primary.mockRestore(); });
  const installed: DefaultFontFace[] = [];
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Unknown Japanese'], 'こんにちは'), Object.assign({
    ...options, fonts: supplied ? [{ family: 'Unknown Japanese', bytes: font.slice(0) }] : [],
  }, {
    [defaultFontInstaller]: async (faces: readonly DefaultFontFace[]) => { installed.push(...faces); },
  }), wire.transport);
  clients.push(client);
  expect(await scriptCalls(manager, 2)).toEqual(expectedCalls(['cjk-jp', false, false], ['cjk-jp', false, false]));
  expect(installed).toMatchObject([{ family: 'Noto Sans JP', kind: 'coverage' }]);
  expect(fallback).toHaveBeenCalledTimes(1);
  expect(primary).toHaveBeenCalledTimes(supplied ? 1 : 0);
  const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
  const runs = list.primitives.filter((primitive) => primitive.kind === 'textBox')
    .flatMap((box) => box.lines.flatMap((line) => line.runs));
  expect(runs.some((run) => run.fontFamily === 'Noto Sans JP' && run.fontId !== 0xffffffff)).toBe(true);
  expect(runs.flatMap((run) => run.glyphs).every((glyph) => glyph.glyphId > 0)).toBe(true);
  expect(runs.some((run) => run.width > 0)).toBe(true);
});

test('missing CJK pack substitutes leave browser rendering and coverage registration unchanged', async () => {
  source({ Calibri: async () => font }, () => async () => font);
  const fallback = spyOn(PptxRenderer.prototype, 'registerFallbackFont');
  restorers.push(() => fallback.mockRestore());
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Calibri'], 'Latin 日本語かな'), options, wire.transport);
  clients.push(client);
  expect(fallback).not.toHaveBeenCalled();
  const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
  const runs = list.primitives.filter((primitive) => primitive.kind === 'textBox')
    .flatMap((box) => box.lines.flatMap((line) => line.runs));
  expect(runs.some((run) => /[日本語かな]/.test(run.text) && run.fontId === 0xffffffff)).toBe(true);
  expect(runs.some((run) => run.text.includes('Latin') && run.fontId !== 0xffffffff)).toBe(true);
});

test('Japanese coverage loads once for an unmapped family and mixed Latin text with matching requested traits', async () => {
  const japanese = await notoSansJP();
  const manager = source({ Calibri: async () => font }, (script) => script === 'cjk-jp' ? async () => japanese : undefined);
  const fallback = spyOn(PptxRenderer.prototype, 'registerFallbackFont');
  restorers.push(() => fallback.mockRestore());
  const wire = connection();
  const client = await createPresentationCommandClient(
    await deck(['BIZ UD明朝', 'Calibri'], 'Latin 日本語かな', 'b="1" i="1"'), options, wire.transport);
  clients.push(client);
  for (const slide of client.state.slides) {
    if (slide.id !== client.state.activeSlideId) await client.dispatch({ kind: 'activateSlide', payload: { slideId: slide.id } });
    const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
    const runs = list.primitives.filter((primitive) => primitive.kind === 'textBox')
      .flatMap((box) => box.lines.flatMap((line) => line.runs));
    const covered = runs.filter((run) => /[日本語かな]/.test(run.text));
    expect(covered.length).toBeGreaterThan(0);
    expect(covered.every((run) => run.fontFamily === 'Noto Sans JP' && run.fontId !== 0xffffffff)).toBe(true);
    expect(covered.flatMap((run) => run.glyphs).every((glyph) => glyph.glyphId > 0)).toBe(true);
  }
  expect(await scriptCalls(manager, 2)).toEqual(expectedCalls(['cjk-jp', false, false], ['cjk-jp', true, true]));
  expect(fallback).toHaveBeenCalledTimes(1);
});

test('four Japanese styles acknowledge one coverage installation and registration', async () => {
  const japanese = await notoSansJP();
  const manager = source({}, (script) => script === 'cjk-jp' ? async () => japanese : undefined);
  const fallback = spyOn(PptxRenderer.prototype, 'registerFallbackFont');
  restorers.push(() => fallback.mockRestore());
  const zip = await JSZip.loadAsync(await deck(['Calibri'], '日本語かな'));
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
  zip.file('ppt/slides/slide1.xml', slide.replace(/<a:r>.*?<\/a:r>/, (run) =>
    ['', 'b="1"', 'i="1"', 'b="1" i="1"'].map((traits) => run.replace('sz="2400"', `sz="2400" ${traits}`)).join('')));
  const install = mock(async (_faces: readonly DefaultFontFace[]) => {});
  const wire = connection();
  const client = await createPresentationCommandClient(await zip.generateAsync({ type: 'uint8array' }), Object.assign({
    ...options, fonts: [{ family: 'Calibri', bytes: font.slice(0) }],
  }, { [defaultFontInstaller]: install }), wire.transport);
  clients.push(client);
  expect(wire.acknowledged).toBe(true);
  expect(await scriptCalls(manager, 5)).toEqual(expectedCalls(
    ['cjk-jp', false, false], ['cjk-jp', false, false], ['cjk-jp', true, false], ['cjk-jp', false, true], ['cjk-jp', true, true],
  ));
  expect(install).toHaveBeenCalledTimes(1);
  expect(install.mock.calls[0][0]).toMatchObject([
    { family: 'Noto Sans JP', bold: false, italic: false, kind: 'coverage' },
  ]);
  expect(fallback).toHaveBeenCalledTimes(1);
});

test('inserting script text loads coverage after the family attempt and repaints the active slide', async () => {
  const japanese = await notoSansJP();
  const requested = deferred<void>();
  const loaded = deferred<ArrayBuffer>();
  const manager = source({ Calibri: async () => font }, (script) => script === 'cjk-jp' ? () => {
    requested.resolve();
    return loaded.promise;
  } : undefined);
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Calibri']), options, wire.transport);
  clients.push(client);
  expect(manager.resolveScriptFallback).not.toHaveBeenCalled();
  const content = (await client.read()).result;
  if (!content.ok) throw new Error(content.failure.message);
  const story = content.stories[0];
  const arrived = deferred<void>();
  client.subscribe((ack) => {
    if ((ack as DefaultFontPayload).defaultFonts?.some((face) => face.kind === 'coverage')) arrived.resolve();
  });
  await client.dispatch({ kind: 'insertText', payload: { inputTime: 1, undoGroup: 'script', text: 'こんにちは ',
    style: { fontFamily: 'Calibri', bold: false, italic: false },
    at: { source: 'confirmed', selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
      anchor: 0, focus: 0, affinity: 'downstream' } } } });
  await requested.promise;
  loaded.resolve(japanese);
  await arrived.promise;
  await client.flush();
  expect(manager.resolve.mock.calls).toEqual([['Calibri', false, false]]);
  expect(manager.resolveScriptFallback.mock.calls).toEqual([['cjk-jp', false, false]]);
  const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
  const runs = list.primitives.filter((primitive) => primitive.kind === 'textBox')
    .flatMap((box) => box.lines.flatMap((line) => line.runs));
  expect(runs.some((run) => /[\u3040-\u309f]/.test(run.text) && run.fontFamily === 'Noto Sans JP')).toBe(true);
});

test('scripts other than CJK keep the browser font without a script default', async () => {
  const carlito = new Uint8Array(await readFile(new URL('../../../fonts/assets/Carlito-Regular.ttf', import.meta.url))).buffer;
  const manager = source({ Calibri: async () => carlito }, () => async () => font);
  const wire = connection();
  const client = await createPresentationCommandClient(await deck(['Calibri'], 'alpha مرحبا שלום नमस्ते ไทย'), options, wire.transport);
  clients.push(client);
  expect(manager.resolveScriptFallback).not.toHaveBeenCalled();
  const list = JSON.parse(new TextDecoder().decode(client.state.frame!.displayList)) as SlideDisplayList;
  const runs = list.primitives.filter((primitive) => primitive.kind === 'textBox')
    .flatMap((box) => box.lines.flatMap((line) => line.runs)).filter((run) => /[\u0590-\u06ff\u0900-\u097f\u0e00-\u0e7f]/.test(run.text));
  expect(runs.length).toBeGreaterThan(0);
  expect(runs.every((run) => run.fontId === 0xffffffff)).toBe(true);
});

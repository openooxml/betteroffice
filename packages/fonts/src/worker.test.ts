import { afterEach, expect, mock, test } from 'bun:test';
import { readFile, readdir } from 'node:fs/promises';
import * as fonts from './index';
import * as worker from './worker';
import * as progressive from './progressive-entry';
import { importFonts, importProgressiveFonts } from './workerFontModules';
import { bundledDecoderUrl, readWasm, setWorkerFontAssets } from './workerAssets';
import { fontAssetResponse } from './test-assets';
import { decodeFontFace, type FontDecoder } from './decode';
import {
  createFontPackResponder, createWorkerFontPacks, type FontAssetKind, type WorkerFontPackSource,
} from '../../../shared/office-session/optionalFontPacks';
import fontsPackage from '../package.json';
import cjk from '../../fonts-cjk/assets/manifest.json';

const originalFetch = globalThis.fetch;
const decoder: FontDecoder = (bytes, options) => decodeFontFace(bytes, options.dict, {
  ...options, wasmUrl: new URL('../dist/decoder/fontpack.wasm', import.meta.url),
});
afterEach(() => {
  setWorkerFontAssets(async () => undefined);
  globalThis.fetch = originalFetch;
});

test('the inlined modules expose the page font and progressive selectors', async () => {
  setWorkerFontAssets(async (pack) => ({ urls: await fonts.fontAssetUrls(pack) }));
  expect(await importFonts()).toBe(worker);
  expect(await importProgressiveFonts()).toBe(progressive);
  for (const name of Object.keys(fonts)) {
    if (name !== 'fontAssetUrls') expect(Object.prototype.hasOwnProperty.call(worker, name)).toBe(true);
  }
});

test('base Latin and the optional Latin pack have separate URL maps', async () => {
  const source = mock(async (pack: FontAssetKind) => ({ urls: await fonts.fontAssetUrls(pack) }));
  setWorkerFontAssets(source);
  const regular = worker.resolveMetricCompatFace('Calibri', false, false)!;
  const [base] = worker.BASE_LATIN_FONTS;
  expect(await Promise.all([worker.bundledFaceInstalled(regular), worker.bundledFaceInstalled(base!)])).toEqual([true, true]);
  expect(source.mock.calls).toEqual([['latin'], ['base']]);
});

test('missing script add-ons load the Latin substitute unless strict loading is requested', async () => {
  const fallback = worker.resolveMetricCompatFace('Arial', false, false)!;
  const source = mock(async (pack: FontAssetKind) => ({
    urls: pack === 'cjk' || pack === 'intl' ? undefined : await fonts.fontAssetUrls(pack),
  }));
  setWorkerFontAssets(source);
  globalThis.fetch = fontAssetResponse as typeof fetch;
  const expected = await fonts.loadBundledFontBytes(fallback, undefined, { decoder });
  for (const [script, name] of [['cjk-sc', 'MissingCjkAddonError'], ['thai', 'MissingIntlAddonError']] as const) {
    const face = worker.resolveScriptFallbackFace(script, false, false)!;
    expect(await worker.bundledFaceInstalled(face)).toBe(false);
    await expect(worker.loadBundledFontBytes(face, undefined, { decoder })).resolves.toEqual(expected);
    const strict = await worker.loadBundledFontBytes(face, undefined, { strict: true }).catch((error: unknown) => error);
    expect(strict).toBeInstanceOf(Error);
    expect((strict as Error).constructor.name).toBe(name);
    expect(fonts.isMissingFontAddon(strict)).toBe(true);
  }
  const location = worker.createFontProvider().resolveScriptFallback('cjk-sc', false, false)!.slices!;
  const sliceError = await location.url(location.ref.index).catch((error: unknown) => error);
  expect((sliceError as Error).constructor.name).toBe('MissingCjkAddonError');
});

test('missing intl and Latin add-ons keep their own errors and are retried', async () => {
  const source = mock<WorkerFontPackSource>(async () => ({ urls: undefined }));
  setWorkerFontAssets(source);
  const intl = worker.resolveScriptFallbackFace('thai', false, false)!;
  const latin = worker.resolveMetricCompatFace('Calibri', false, false)!;
  const intlError = await worker.loadBundledFontBytes(intl, undefined, { strict: true }).catch((error: unknown) => error);
  expect((intlError as Error).constructor.name).toBe('MissingIntlAddonError');
  await expect(worker.loadBundledFontBytes(latin)).rejects.toMatchObject({ name: 'MissingFontAddonError' });
  await expect(worker.bundledFaceInstalled(intl)).resolves.toBe(false);
  expect(source.mock.calls).toEqual([['intl'], ['latin'], ['intl']]);
});

test('absolute and page-relative base URLs override the pack source for worker faces and slices', async () => {
  let pageUrl = 'https://page.example/documents/current';
  const source = mock<WorkerFontPackSource>(async (kind) => {
    expect(kind).toBe('base');
    return { pageUrl };
  });
  setWorkerFontAssets(source);
  const requests: string[] = [];
  globalThis.fetch = ((input: RequestInfo | URL) => {
    requests.push(String(input));
    return fontAssetResponse(input);
  }) as typeof fetch;
  const face = worker.resolveMetricCompatFace('Calibri', false, false)!;
  const absolute = new URL('https://assets.example/custom');
  expect((await worker.loadBundledFontBytes(face, { baseUrl: absolute }, { decoder })).byteLength).toBe(face.byteLength);
  expect(source).not.toHaveBeenCalled();
  const provider = worker.createFontProvider({ baseUrl: './custom' }, { decoder });
  expect(source).toHaveBeenCalledTimes(1);
  pageUrl = 'https://page.example/another/route';
  expect((await provider.resolve('Calibri', false, false)!()).byteLength).toBe(face.byteLength);
  const slices = provider.resolveScriptFallback('cjk-jp', false, false)!.slices!;
  expect((await slices.url(slices.ref.index)).href).toBe(`https://page.example/documents/custom/${slices.ref.index}`);
  expect((await slices.prepareUrls!()).base).toBe('https://page.example/documents/custom/');
  expect(requests).toContain(`https://assets.example/custom/${face.file}`);
  expect(requests).toContain(`https://page.example/documents/custom/${face.file}`);
  expect(source).toHaveBeenCalledTimes(1);
});

test('providers created across navigation pin separate page URLs while a base pack lookup is pending', async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'location');
  const location = { href: 'https://page.example/first/document' };
  Object.defineProperty(globalThis, 'location', { configurable: true, value: location });
  const channel = createWorkerFontPacks((message) => { page.receive(message); });
  const load = mock(async () => ({}));
  const page = createFontPackResponder(channel.receive, load);
  setWorkerFontAssets(channel.source);
  try {
    const first = worker.createFontProvider({ baseUrl: './custom' });
    location.href = 'https://page.example/second/document';
    const second = worker.createFontProvider({ baseUrl: './custom' });
    location.href = 'https://page.example/third/document';
    const sliceBase = async (provider: ReturnType<typeof worker.createFontProvider>) => {
      const slices = provider.resolveScriptFallback('cjk-jp', false, false)!.slices!;
      return (await slices.prepareUrls!()).base;
    };
    expect(await Promise.all([sliceBase(first), sliceBase(second)])).toEqual([
      'https://page.example/first/custom/', 'https://page.example/second/custom/',
    ]);
    expect(load).toHaveBeenCalledTimes(1);
  } finally {
    channel.dispose();
    if (original) Object.defineProperty(globalThis, 'location', original);
    else Reflect.deleteProperty(globalThis, 'location');
  }
});

test('a relative font root without a page URL rejects even with a worker script location', async () => {
  const location = Object.getOwnPropertyDescriptor(globalThis, 'location');
  Object.defineProperty(globalThis, 'location', { configurable: true, value: { href: 'https://worker.example/script/worker.js' } });
  try {
    for (const response of [{ urls: {} }, undefined]) {
      setWorkerFontAssets(async () => response);
      const options = { baseUrl: './custom' };
      const provider = worker.createFontProvider(options, { decoder });
      const face = worker.resolveMetricCompatFace('Calibri', false, false)!;
      const slices = provider.resolveScriptFallback('cjk-jp', false, false)!.slices!;
      for (const load of [provider.resolve('Calibri', false, false)!, () => slices.url(slices.ref.index),
        () => worker.loadBundledFontBytes(face, options, { decoder })]) {
        const error = await load().catch((reason: unknown) => reason);
        expect(error).toBeInstanceOf(TypeError);
        expect(error).toMatchObject({ message: 'Font baseUrl must be absolute when no browser location exists: ./custom/' });
      }
    }
  } finally {
    if (location) Object.defineProperty(globalThis, 'location', location);
    else Reflect.deleteProperty(globalThis, 'location');
  }
});

test('the decoder URL comes from its own request and retries misses', async () => {
  const href = 'https://page.example/assets/published-decoder.wasm';
  const source = mock<WorkerFontPackSource>(async () => ({ urls: { 'fontpack.wasm': href } })).mockResolvedValueOnce({});
  setWorkerFontAssets(source);
  expect(await bundledDecoderUrl()).toBeUndefined();
  expect((await bundledDecoderUrl())?.href).toBe(href);
  expect(source.mock.calls).toEqual([['decoder'], ['decoder']]);
});

test('the worker decoder preserves file and HTTP missing-asset details', async () => {
  const fetchWasm = mock(async () => new Response(null, { status: 404 }));
  globalThis.fetch = fetchWasm as unknown as typeof fetch;
  await expect(readWasm(new URL('./missing-decoder/fontpack.wasm', import.meta.url))).rejects.toMatchObject({ code: 'ENOENT' });
  expect(fetchWasm).not.toHaveBeenCalled();
  await expect(readWasm(new URL('https://fonts.example/fontpack.wasm'))).rejects.toMatchObject({ decoderStatus: 404 });
});

test('the worker decoder recognizes Deno NotFound without Node file APIs', async () => {
  const host = globalThis.process as { getBuiltinModule?: unknown };
  const original = host.getBuiltinModule;
  const deno = Object.getOwnPropertyDescriptor(globalThis, 'Deno');
  const readFile = mock(async () => { throw Object.assign(new Error('No such file'), { name: 'NotFound' }); });
  host.getBuiltinModule = undefined;
  Object.defineProperty(globalThis, 'Deno', { value: { readFile }, configurable: true });
  try {
    const url = new URL('./missing-deno-decoder/fontpack.wasm', import.meta.url);
    await expect(readWasm(url)).rejects.toMatchObject({ code: 'ENOENT' });
    expect(readFile).toHaveBeenCalledWith(url);
  } finally {
    host.getBuiltinModule = original;
    if (deno) Object.defineProperty(globalThis, 'Deno', deno);
    else Reflect.deleteProperty(globalThis, 'Deno');
  }
});

test('the published main entry leaves all five slice manifests in the CJK add-on', async () => {
  expect(Object.keys(fontsPackage.exports)).not.toContain('./worker');
  const types = await readFile(new URL('../dist/index.d.ts', import.meta.url), 'utf8');
  expect(types).toContain('fontAssetUrls');
  const bundle = await readFile(new URL('../dist/index.js', import.meta.url), 'utf8');
  const main: typeof fonts = await import(new URL('../dist/index.js', import.meta.url).href);
  expect(await main.fontAssetUrls('decoder')).toEqual({ 'fontpack.wasm': new URL('../dist/decoder/fontpack.wasm', import.meta.url).href });
  const urls = await main.fontAssetUrls('cjk');
  const manifests = cjk.faces.flatMap(({ slices }) => slices ? [slices.index] : []);
  expect(manifests).toHaveLength(5);
  for (const file of manifests) {
    const index = JSON.parse(await readFile(new URL(urls![file]!), 'utf8')) as { chunks: { sha256: string }[] };
    expect(bundle).not.toContain(index.chunks[0]!.sha256);
  }
});

test('published font declarations keep CDN session hooks internal', async () => {
  const dist = new URL('../dist/', import.meta.url);
  const entries = (await readdir(dist)).filter((file) => file.endsWith('.d.ts'));
  expect(entries).toEqual(expect.arrayContaining(['index.d.ts', 'progressive.d.ts']));
  for (const file of entries) {
    const types = await readFile(new URL(file, dist), 'utf8');
    expect(types).not.toMatch(/\b(?:FontCdnReporter|reportCdnFailure|checkImplicitCdnPack|onSubstitute)\b/);
  }
});

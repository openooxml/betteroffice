import { afterEach, beforeEach, expect, jest, mock, spyOn, test } from 'bun:test';
import * as Bun from 'bun';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Worker } from 'node:worker_threads';
import latin from '../../fonts-latin/assets/manifest.json';
import intl from '../../fonts-intl/assets/manifest.json';
import cjk from '../../fonts-cjk/assets/manifest.json';
import { decodeFontFace, type FontDecoder } from './decode';
import * as fontCdn from './fontCdn';
import { createFontRuntime, type FontPackDiagnostic } from './fontRuntime';
import { createPageFontPackResponder } from '../../../shared/office-session/fontPackSource';
import { fontAssetResponse, withBrowserFontGlobals } from './fontAssets.test-support';
import { openSlicedFace, sliceSourceOf } from './progressive-entry';
import { parseSkeleton } from './slices';
import { writeFontCdnTree } from '../../../scripts/fonts/cdn-tree';
import { getFontCdnConfiguration, implicitFontCdnBases, setFontCdnConfiguration, type FontCdnConfiguration } from '../../../shared/office-session/fontCdnRuntime';
import { createFontPackResponder, createWorkerFontPacks } from '../../../shared/office-session/optionalFontPacks';
import { withinTestBudget } from '../../../shared/office-session/testing/loadBudget';
import { setWorkerFontAssets } from './workerAssets';
import * as worker from './worker';

const { fontCdnAssetUrl } = fontCdn;

let installed = false;
for (const [name, manifest, exported] of [
  ['latin', latin, 'LATIN_FONT_ASSET_URLS'],
  ['intl', intl, 'INTL_FONT_ASSET_URLS'],
  ['cjk', cjk, 'CJK_FONT_ASSET_URLS'],
] as const) {
  mock.module(`@betteroffice/fonts-${name}`, () => ({
    then(resolve: (addon: unknown) => void, reject: (error: unknown) => void) {
      Promise.resolve().then(async () => {
        if (!installed) throw new Error(`Cannot find package '@betteroffice/fonts-${name}'`);
        const files = manifest.faces.map(({ file }) => file);
        if (name === 'cjk') {
          for (const face of cjk.faces) {
            if (!face.slices) continue;
            const index = JSON.parse(await Bun.file(new URL(`../../fonts-cjk/assets/${face.slices.index}`, import.meta.url)).text()) as {
              skeleton: { file: string }; core: { file: string }; chunks: { file: string }[];
            };
            files.push(face.slices.index, index.skeleton.file, index.core.file, ...index.chunks.map(({ file }) => file));
          }
        }
        return { [exported]: Object.fromEntries(files.map((file) => [file,
          () => new URL(`../../fonts-${name}/assets/${file}`, import.meta.url)])) };
      }).then(resolve, reject);
    },
  }));
}

const fonts = await import('./index');
const originalFetch = globalThis.fetch;
const decoder: FontDecoder = (bytes, options) => decodeFontFace(bytes, options.dict, options);
const hash = (bytes: ArrayBuffer) => createHash('sha256').update(new Uint8Array(bytes)).digest('hex');
const regular = fonts.resolveMetricCompatFace('Calibri', false, false)!;
const japanese = fonts.resolveScriptFallbackFace('cjk-jp', false, false)!;
const runtime = { decoder, retain: false, strict: true };
const asset = (pack: string, file: string) => new URL(`../../fonts-${pack}/assets/${file}`, import.meta.url);

function serveAssets(requests: string[], inspect?: (url: URL, options?: RequestInit) => void): void {
  globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
    const url = new URL(String(input));
    if (url.pathname.endsWith('.wasm')) return fontAssetResponse(input);
    requests.push(url.href);
    inspect?.(url, options);
    if (url.protocol === 'file:') return new Response(await Bun.file(url).arrayBuffer());
    const match = url.pathname.match(/\/(latin|intl|cjk)\/[^/]+\/(.+)$/) ??
      url.pathname.match(/\/fonts-(latin|intl|cjk)@[^/]+\/assets\/(.+)$/);
    if (!match) return fontAssetResponse(input);
    return new Response(await Bun.file(asset(match[1]!, match[2]!)).arrayBuffer());
  }) as typeof fetch;
}

let previousConfiguration: FontCdnConfiguration | undefined;
beforeEach(() => {
  previousConfiguration = getFontCdnConfiguration();
  setFontCdnConfiguration(undefined);
  implicitFontCdnBases().clear();
});

afterEach(() => {
  fonts.configureFontCdn({ baseUrl: false });
  setFontCdnConfiguration(previousConfiguration);
  setWorkerFontAssets(async () => undefined);
  globalThis.fetch = originalFetch;
});

test('CDN loading starts disabled and preserves missing-pack errors', async () => {
  expect(getFontCdnConfiguration()).toBeUndefined();
  expect(fonts.fontCdnConfiguration()).toEqual({ baseUrl: false, timeoutMs: 30_000 });
  expect(getFontCdnConfiguration()).toBeUndefined();
  expect(await fonts.bundledFaceInstalled(regular)).toBe(false);
  await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toMatchObject({ name: 'MissingFontAddonError' });
  setWorkerFontAssets(async () => ({ urls: undefined }));
  for (const face of [japanese, fonts.resolveScriptFallbackFace('thai', false, false)!, regular]) {
    const pageError = await fonts.loadBundledFontBytes(face, undefined, runtime).catch((error: unknown) => error);
    const workerError = await worker.loadBundledFontBytes(face, undefined, runtime).catch((error: unknown) => error);
    expect(pageError).toBeInstanceOf(Error);
    expect((workerError as Error).constructor).toBe((pageError as Error).constructor);
    expect((workerError as Error).message).toBe((pageError as Error).message);
  }
});

test('the worker CDN overrides the page CDN, including when explicitly disabled', async () => {
  const pageCdn = { baseUrl: 'https://page-fonts.example/', timeoutMs: 4321 };
  const channel = createWorkerFontPacks((message) => { page.receive(message); });
  const page = createFontPackResponder(channel.receive, async () => undefined, () => pageCdn);
  setWorkerFontAssets(channel.source);
  const requests: string[] = [];
  serveAssets(requests);
  try {
    await worker.loadBundledFontBytes(regular, undefined, runtime);
    worker.configureFontCdn({ baseUrl: 'https://worker-fonts.example/', timeoutMs: 1234 });
    await worker.loadBundledFontBytes(regular, undefined, runtime);
    expect(requests).toEqual([
      `https://page-fonts.example/latin/${latin.version}/${regular.file}`,
      `https://worker-fonts.example/latin/${latin.version}/${regular.file}`,
    ]);
    const slices = worker.createFontProvider().resolveScriptFallback('cjk-jp', false, false)!.slices!;
    const prepared = await slices.prepareUrls!();
    expect(prepared.timeoutMs).toBe(1234);
    expect((await prepared.url(slices.ref.index)).href).toBe(`https://worker-fonts.example/cjk/${cjk.version}/${slices.ref.index}`);
    worker.configureFontCdn({ baseUrl: false });
    await expect(worker.bundledFaceInstalled(regular)).resolves.toBe(false);
  } finally {
    channel.dispose();
  }
});

test('a worker sees CDN changes after a successful pack lookup and resolves CJK slices from the published index', async () => {
  const channel = createWorkerFontPacks((message) => { page.receive(message); });
  const page = createFontPackResponder(channel.receive, fonts.fontAssetUrls, getFontCdnConfiguration);
  setWorkerFontAssets(channel.source);
  const requests: string[] = [];
  serveAssets(requests);
  const provider = worker.createFontProvider(undefined, { decoder });
  const location = provider.resolveScriptFallback('cjk-jp', false, false)!.slices!;
  try {
    fonts.configureFontCdn({ baseUrl: 'https://fonts.example/first', timeoutMs: 1234 });
    expect(await worker.bundledFaceInstalled(japanese)).toBe(true);
    const first = await openSlicedFace(location);
    expect(requests).toContain(`https://fonts.example/first/cjk/${cjk.version}/${location.ref.index}`);
    expect((await sliceSourceOf(first)!).timeoutMs).toBe(1234);
    fonts.configureFontCdn({ baseUrl: 'https://fonts.example/second', timeoutMs: 4321 });
    const second = await openSlicedFace(location);
    await second.ensureText('日本語');
    expect(requests).toContain(`https://fonts.example/second/cjk/${cjk.version}/${location.ref.index}`);
    const source = (await sliceSourceOf(second)!)!;
    expect(source.timeoutMs).toBe(4321);
    expect(Object.keys(source.urls).some((file) => /\/c\d+\.zst$/.test(file))).toBe(true);
    expect(Object.values(source.urls).every((url) => url.startsWith('https://fonts.example/second/'))).toBe(true);
    fonts.configureFontCdn({ baseUrl: false });
    expect(await worker.bundledFaceInstalled(japanese)).toBe(false);
  } finally {
    channel.dispose();
  }
});

test('an absent Latin add-on has no map or diagnostic even with built-in Latin and a configured CDN', async () => {
  const channel = createWorkerFontPacks((message) => { page.receive(message); });
  const page = createFontPackResponder(channel.receive, fonts.fontAssetUrls, getFontCdnConfiguration);
  setWorkerFontAssets(channel.source);
  try {
    for (const baseUrl of [false, 'https://fonts.example/'] as const) {
      fonts.configureFontCdn({ baseUrl });
      expect(await fonts.fontAssetUrls('latin')).toBeUndefined();
      const base = await fonts.fontAssetUrls('base');
      expect(Object.keys(base!).sort()).toEqual(fonts.BASE_LATIN_FONTS.map(({ file }) => file).sort());
      expect(await worker.bundledFaceInstalled(fonts.BASE_LATIN_FONTS[0]!)).toBe(true);
      expect(await worker.bundledFaceInstalled(regular)).toBe(Boolean(baseUrl));
      expect(await worker.checkFontPacks()).toEqual([]);
    }
  } finally {
    channel.dispose();
  }
});

test('session clients and separate package copies share font CDN configuration', async () => {
  fonts.configureFontCdn({ baseUrl: new URL('https://shared-fonts.example/root'), timeoutMs: 4321 });
  const configured = { baseUrl: 'https://shared-fonts.example/root/', timeoutMs: 4321 };
  expect(Reflect.get(globalThis, Symbol.for('betteroffice.fonts.cdn'))).toEqual(configured);
  expect(getFontCdnConfiguration()).toEqual(configured);
  const copy: typeof import('./fontCdn') = await import(new URL('./fontCdn.ts?copy', import.meta.url).href);
  expect(copy.fontCdnConfiguration()).toEqual(configured);
  copy.configureFontCdn({ baseUrl: false, timeoutMs: 1234 });
  expect(fonts.fontCdnConfiguration()).toEqual({ baseUrl: false, timeoutMs: 1234 });
  expect(getFontCdnConfiguration()).toEqual({ baseUrl: false, timeoutMs: 1234 });
  fonts.configureFontCdn({ baseUrl: configured.baseUrl });
  expect(copy.fontCdnConfiguration()).toEqual({ baseUrl: configured.baseUrl, timeoutMs: 30_000 });
  const snapshot = getFontCdnConfiguration()!;
  snapshot.baseUrl = false;
  expect(getFontCdnConfiguration()).toEqual({ baseUrl: configured.baseUrl, timeoutMs: 30_000 });
});

test('a missing Latin pack loads pinned bytes with private requests and a package decoder', async () => {
  const requests: string[] = [];
  const decoderRequests: string[] = [];
  fonts.configureFontCdn({ baseUrl: new URL('https://configured.example/fonts') });
  serveAssets(requests, (_url, options) => {
    expect(options?.credentials).toBe('omit');
    expect(options?.referrerPolicy).toBe('no-referrer');
    expect(options?.signal).toBeInstanceOf(AbortSignal);
  });
  const fetchAsset = globalThis.fetch;
  globalThis.fetch = ((input: RequestInfo | URL, options?: RequestInit) => {
    if (String(input).endsWith('.wasm')) decoderRequests.push(String(input));
    return fetchAsset(input, options);
  }) as typeof fetch;
  expect(await fonts.bundledFaceInstalled(regular)).toBe(true);
  expect(requests).toEqual([]);
  const bytes = await fonts.loadBundledFontBytes(regular);
  expect(bytes.byteLength).toBe(regular.byteLength);
  expect(hash(bytes)).toBe(latin.faces.find(({ file }) => file === regular.file)!.sha256);
  expect(requests).toEqual([`https://configured.example/fonts/latin/${latin.version}/${regular.file}`]);
  expect(decoderRequests.every((url) => !url.startsWith('https://configured.example/'))).toBe(true);
});

test('an explicit asset root keeps its existing request policy', async () => {
  const requests: string[] = [];
  fonts.configureFontCdn({ baseUrl: 'https://fallback.example/' });
  serveAssets(requests, (_url, options) => expect(options).toBeUndefined());
  await fonts.loadBundledFontBytes(regular, { baseUrl: 'https://explicit.example/' }, runtime);
  expect(requests).toEqual([`https://explicit.example/${regular.file}`]);
});

test('base Latin stays in the package when a CDN is configured', async () => {
  const requests: string[] = [];
  fonts.configureFontCdn({ baseUrl: 'https://base-latin.example/' });
  serveAssets(requests);
  const bytes = await fonts.createFontProvider(undefined, runtime).resolveBaseLatin!(false, false)!();
  expect(bytes.byteLength).toBeGreaterThan(0);
  expect(requests.every((url) => !url.startsWith('https://base-latin.example/'))).toBe(true);
});

test('a dictionary prerequisite uses the same pack version directory', async () => {
  const requests: string[] = [];
  const face = fonts.resolveMetricCompatFace('Calibri', true, false)!;
  const entry = latin.faces.find(({ file }) => file === face.file)!;
  fonts.configureFontCdn({ baseUrl: 'https://dictionary.example/root/' });
  serveAssets(requests);
  const bytes = await fonts.loadBundledFontBytes(face, undefined, runtime);
  expect(hash(bytes)).toBe(entry.sha256);
  expect(entry.dict).not.toBeNull();
  expect(requests.sort()).toEqual([entry.file, entry.dict].map((file) =>
    `https://dictionary.example/root/latin/${latin.version}/${file}`).sort());
});

test('404 and timeout failures are evicted so the same request can succeed', async () => {
  for (const failure of ['404', 'timeout']) {
    fonts.configureFontCdn({ baseUrl: `https://${failure}.example/`, timeoutMs: 10 });
    const clock = cdnClock();
    const started = deferred<void>();
    let calls = 0;
    globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
      if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
      calls++;
      if (calls === 1) {
        if (failure === '404') return new Response(null, { status: 404 });
        started.resolve();
        return waitForAbort(AbortSignal.any([options!.signal!, clock.signal]));
      }
      return new Response(await Bun.file(asset('latin', regular.file)).arrayBuffer());
    }) as typeof fetch;
    try {
      const result = fonts.loadBundledFontBytes(regular, undefined, { decoder, strict: true });
      if (failure === 'timeout') {
        await beforeFontLoadFinishes(started.promise, result, 'the first CDN request');
        await clock.advance(10);
      }
      await expect(withinTestBudget(result, 'the CDN failure')).rejects.toThrow(failure === '404' ? 'HTTP 404' : 'Timed out');
      expect((await fonts.loadBundledFontBytes(regular, undefined, { decoder, strict: true })).byteLength).toBe(regular.byteLength);
      expect(calls).toBe(2);
    } finally {
      clock.restore();
    }
  }
});

test('CDN faces enforce encoded length, decoded length and SHA-256 before retrying', async () => {
  fonts.configureFontCdn({ baseUrl: 'https://integrity.example/' });
  let mode = 'encoded';
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
    return new Response(mode === 'encoded' ? new Uint8Array(3) : await Bun.file(asset('latin', regular.file)).arrayBuffer());
  }) as typeof fetch;
  const checked: FontDecoder = async (bytes, options) => {
    if (mode === 'decoded') return { font: new ArrayBuffer(1), payload: new ArrayBuffer(1) };
    if (mode === 'hash') return { font: new ArrayBuffer(regular.byteLength), payload: new ArrayBuffer(1) };
    return decoder(bytes, options);
  };
  for (const [next, error] of [['encoded', 'encoded bytes'], ['decoded', 'bytes; expected'], ['hash', 'SHA-256 mismatch']]) {
    mode = next!;
    await expect(fonts.loadBundledFontBytes(regular, undefined, { decoder: checked })).rejects.toThrow(error);
  }
  mode = 'valid';
  expect((await fonts.loadBundledFontBytes(regular, undefined, { decoder: checked })).byteLength).toBe(regular.byteLength);
});

test('a face without manifest integrity metadata cannot take a CDN path', async () => {
  fonts.configureFontCdn({ baseUrl: 'https://unmanifested.example/' });
  const requests: string[] = [];
  serveAssets(requests);
  const face = { ...japanese, file: 'Synthetic-Regular.ttf', byteLength: 4 };
  await expect(fonts.loadBundledFontBytes(face, undefined, runtime)).rejects.toThrow();
  expect(requests).toEqual([]);
});

test('CJK indexes, skeletons and chunks use the CJK version and preserve worker transport policy', async () => {
  const requests: string[] = [];
  fonts.configureFontCdn({ baseUrl: 'https://sliced.example/', timeoutMs: 1234 });
  serveAssets(requests, (_url, options) => {
    expect(options?.credentials).toBe('omit');
    expect(options?.referrerPolicy).toBe('no-referrer');
  });
  const location = fonts.createFontProvider(undefined, runtime).resolveScriptFallback('cjk-jp', false, false)!.slices!;
  const face = await openSlicedFace(location);
  const entry = cjk.faces.find(({ file }) => file === japanese.file)!;
  const index = JSON.parse(await Bun.file(asset('cjk', entry.slices!.index)).text()) as { skeleton: { file: string } };
  const encoded = await Bun.file(asset('cjk', index.skeleton.file)).arrayBuffer();
  const skeleton = parseSkeleton(new Uint8Array((await decodeFontFace(encoded)).font));
  let codepoint = 0x4e00;
  while (codepoint <= 0x9fff && skeleton.chunkOfCodepoint(codepoint) !== 1) codepoint++;
  expect(codepoint).toBeLessThanOrEqual(0x9fff);
  await face.ensureText(String.fromCodePoint(codepoint));
  expect(requests).toContain(`https://sliced.example/cjk/${cjk.version}/${location.ref.index}`);
  expect(requests.some((url) => /\/c\d+\.zst$/.test(url))).toBe(true);
  expect(requests.every((url) => url.startsWith(`https://sliced.example/cjk/${cjk.version}/`))).toBe(true);
  const source = await sliceSourceOf(face)!;
  expect(source.timeoutMs).toBe(1234);
  expect(Object.values(source.urls).every((url) => url.startsWith(`https://sliced.example/cjk/${cjk.version}/`))).toBe(true);
});

test('every sliced CDN file validates its bytes and failed chunks can retry', async () => {
  const entry = cjk.faces.find(({ file }) => file === japanese.file)!;
  const index = JSON.parse(await Bun.file(asset('cjk', entry.slices!.index)).text()) as {
    skeleton: { file: string }; core: { file: string }; chunks: { file: string }[];
  };
  globalThis.fetch = fontAssetResponse as typeof fetch;
  const skeleton = parseSkeleton(new Uint8Array((await decodeFontFace(await Bun.file(asset('cjk', index.skeleton.file)).arrayBuffer())).font));
  let chunkGlyph = 0;
  while (chunkGlyph < skeleton.numGlyphs && skeleton.chunkOfGlyph(chunkGlyph) !== 1) chunkGlyph++;
  expect(chunkGlyph).toBeLessThan(skeleton.numGlyphs);
  for (const target of [entry.slices!.index, index.skeleton.file, index.core.file, index.chunks[0]!.file]) {
    fonts.configureFontCdn({ baseUrl: `https://sliced-integrity.example/${encodeURIComponent(target)}/` });
    const poisoned = new WeakSet<object>();
    let tamper = true;
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
      const file = new URL(String(input)).pathname.split(`/cjk/${cjk.version}/`)[1]!;
      const bytes = new Uint8Array(await Bun.file(asset('cjk', file)).arrayBuffer());
      if (tamper && file === target && file.endsWith('.json')) bytes[0] ^= 1;
      return new Response(bytes);
    }) as typeof fetch;
    const checked: FontDecoder = async (encoded, options) => {
      const response = await decoder(encoded, options);
      if (tamper && poisoned.has(encoded)) new Uint8Array(response.font)[0] ^= 1;
      return response;
    };
    const fetchAsset = globalThis.fetch;
    globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
      const response = await fetchAsset(input, options);
      if (new URL(String(input)).pathname.endsWith(target) && !target.endsWith('.json')) {
        const read = response.arrayBuffer.bind(response);
        response.arrayBuffer = async () => {
          const bytes = await read();
          poisoned.add(bytes);
          return bytes;
        };
      }
      return response;
    }) as typeof fetch;
    const location = fonts.createFontProvider(undefined, { decoder: checked, retain: false }).resolveScriptFallback('cjk-jp', false, false)!.slices!;
    if (target === index.chunks[0]!.file) {
      const face = await openSlicedFace(location);
      await expect(face.ensureGlyphs([chunkGlyph])).rejects.toThrow('SHA-256 mismatch');
      tamper = false;
      await expect(face.ensureGlyphs([chunkGlyph])).resolves.toBeGreaterThan(1);
    } else {
      await expect(openSlicedFace(location)).rejects.toThrow('SHA-256 mismatch');
      tamper = false;
      await expect(openSlicedFace(location)).resolves.toBeDefined();
    }
  }
});

test('the future jsDelivr mapper matches the existing CDN entry and uses npm versions', async () => {
  const { createFontProvider } = await import('./cdn');
  const requests: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
    requests.push(String(input));
    return fontAssetResponse(input);
  }) as typeof fetch;
  const provider = createFontProvider(undefined, { decoder, retain: false });
  await provider.resolve('Calibri', false, false)!();
  await provider.resolveScriptFallback('thai', false, false)!();
  const slices = provider.resolveScriptFallback('cjk-jp', false, false)!.slices!;
  for (const [pack, file] of [[latin, regular.file], [intl, fonts.resolveScriptFallbackFace('thai', false, false)!.file]] as const) {
    expect(requests).toContain(fontCdnAssetUrl({ ...pack, version: 'manifest-version-is-unused' }, file).href);
  }
  expect((await slices.url(slices.ref.index)).href).toBe(fontCdnAssetUrl(cjk, slices.ref.index).href);
});

test('the CDN tree serves Latin, intl and sliced CJK faces from a local origin', async () => {
  const output = await mkdtemp(join(tmpdir(), 'font-cdn-tree-'));
  const requests: string[] = [];
  let server: ReturnType<typeof Bun.serve> | undefined;
  try {
    await writeFontCdnTree(output);
    server = Bun.serve({
      hostname: '127.0.0.1', port: 0,
      async fetch(request) {
        const file = new URL(request.url).pathname.slice(1);
        requests.push(file);
        return new Response(await readFile(join(output, file)));
      },
    });
    fonts.configureFontCdn({ baseUrl: server.url });
    globalThis.fetch = ((input: RequestInfo | URL, options?: RequestInit) =>
      String(input).endsWith('.wasm') ? fontAssetResponse(input) : originalFetch(input, options)) as typeof fetch;
    const provider = fonts.createFontProvider(undefined, runtime);
    expect((await provider.resolve('Calibri', false, false)!()).byteLength).toBe(regular.byteLength);
    const thai = fonts.resolveScriptFallbackFace('thai', false, false)!;
    expect((await provider.resolveScriptFallback('thai', false, false)!()).byteLength).toBe(thai.byteLength);
    const location = provider.resolveScriptFallback('cjk-jp', false, false)!.slices!;
    const face = await openSlicedFace(location);
    await face.ensureText('日本語');
    expect(requests).toContain(`latin/${latin.version}/${regular.file}`);
    expect(requests).toContain(`intl/${intl.version}/${thai.file}`);
    expect(requests).toContain(`cjk/${cjk.version}/${location.ref.index}`);
    for (const pack of [latin, intl, cjk]) {
      expect(await Bun.file(join(output, pack.pack, pack.version, 'manifest.json')).exists()).toBe(true);
    }
    expect(await Bun.file(join(output, 'latin', latin.version, 'decoder/fontpack.wasm')).exists()).toBe(false);
  } finally {
    server?.stop(true);
    await rm(output, { recursive: true, force: true });
  }
}, 60_000);

const jsDelivr = 'https://cdn.jsdelivr.net/npm/';
const implicitCdn = { baseUrl: jsDelivr, implicit: true } as const;
const stallMs = 10_000;
const cooldownMs = 60_000;
const regularHash = latin.faces.find(({ file }) => file === regular.file)!.sha256;
const regularUrl = fontCdnAssetUrl(latin, regular.file).href;

function observeCdnFailures() {
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const diagnostics: FontPackDiagnostic[] = [];
  const dispatch = Object.getOwnPropertyDescriptor(globalThis, 'dispatchEvent');
  Object.defineProperty(globalThis, 'dispatchEvent', {
    configurable: true,
    value(event: CustomEvent<FontPackDiagnostic>) {
      if (event.type === 'betteroffice:font-pack-diagnostic') diagnostics.push(event.detail);
      return true;
    },
  });
  return {
    warning,
    diagnostics,
    restore() {
      if (dispatch) Object.defineProperty(globalThis, 'dispatchEvent', dispatch);
      else Reflect.deleteProperty(globalThis, 'dispatchEvent');
      warning.mockRestore();
    },
  };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((complete) => { resolve = complete; });
  return { promise, resolve };
}

function beforeFontLoadFinishes<T>(event: Promise<T>, load: Promise<unknown>, description: string): Promise<T> {
  return withinTestBudget(Promise.race([event, load.then(() => {
    throw new Error(`Font load finished before ${description}`);
  })]), description);
}

function cdnClock() {
  jest.useFakeTimers();
  const cleanup = new AbortController();
  let now = Date.now();
  const date = spyOn(Date, 'now').mockImplementation(() => now);
  const totalTimeout = spyOn(AbortSignal, 'timeout').mockImplementation((ms) => {
    const controller = new AbortController();
    setTimeout(() => controller.abort(new DOMException('Timed out', 'TimeoutError')), ms);
    return controller.signal;
  });
  return {
    signal: cleanup.signal,
    async advance(ms: number) {
      now += ms;
      jest.advanceTimersByTime(ms);
      await Bun.sleep(0);
    },
    restore() {
      try {
        cleanup.abort();
        jest.clearAllTimers();
      } finally {
        date.mockRestore();
        totalTimeout.mockRestore();
        jest.useRealTimers();
      }
    },
  };
}

function waitForAbort(signal: AbortSignal): Promise<never> {
  return new Promise((_resolve, reject) => {
    if (signal.aborted) reject(signal.reason);
    else signal.addEventListener('abort', () => reject(signal.reason), { once: true });
  });
}

function streamed(bytes: Uint8Array, signal: AbortSignal, pulls: readonly ReturnType<typeof deferred<void>>[]): Response {
  const midpoint = Math.ceil(bytes.length / 2);
  let offset = 0;
  let index = 0;
  return new Response(new ReadableStream<Uint8Array>({
    async pull(controller) {
      signal.throwIfAborted();
      await new Promise<void>((resolve, reject) => {
        const timer = setTimeout(() => {
          signal.removeEventListener('abort', abort);
          resolve();
        }, 6_000);
        const abort = () => {
          clearTimeout(timer);
          reject(signal.reason);
        };
        signal.addEventListener('abort', abort, { once: true });
        pulls[index++]!.resolve();
      });
      const end = Math.min(offset + midpoint, bytes.length);
      controller.enqueue(bytes.subarray(offset, end));
      offset = end;
      if (offset === bytes.length) controller.close();
    },
  }));
}

test('only a real browser page decides the default CDN', () => {
  const document = {};
  const dom = { window: { document }, document };
  const node = { versions: { node: '22.0.0' } };
  const bun = { versions: { ...node.versions, bun: '1.0.0' } };
  const electron = { versions: { ...node.versions, electron: '33.0.0' } };
  const detect = Reflect.get(fontCdn, 'isBrowserFontPage') as ((globals: object) => boolean) | undefined;
  for (const [name, globals, expected] of [
    ['page', dom, true],
    ['dedicated worker', { importScripts() {} }, false],
    ['module worker', { postMessage() {} }, false],
    ['jsdom under Node', { ...dom, process: node }, false],
    ['happy-dom under Bun', { ...dom, process: bun }, false],
    ['Bun without a process global', { ...dom, Bun: {} }, false],
    ['Electron renderer with Node integration', { ...dom, process: electron }, false],
    ['Electron worker with Node integration', { process: electron, postMessage() {} }, false],
    ['Deno', { ...dom, Deno: {} }, false],
    ['Bun', { process: bun }, false],
    ['Node main', { process: node }, false],
    ['Node worker_threads', { process: node, postMessage() {} }, false],
  ] as const) {
    expect({ name, page: detect?.(globals) ?? false }).toEqual({ name, page: expected });
  }
});

test('browser default serves pinned, checked bytes for an absent pack', async () => {
  const restore = withBrowserFontGlobals();
  const requests: string[] = [];
  serveAssets(requests);
  try {
    const bytes = await fonts.loadBundledFontBytes(regular, undefined, runtime).catch(() => new ArrayBuffer(0));
    expect(bytes.byteLength).toBe(regular.byteLength);
    expect(hash(bytes)).toBe(latin.faces.find(({ file }) => file === regular.file)!.sha256);
    expect(requests).toEqual([fontCdnAssetUrl(latin, regular.file).href]);
  } finally { restore(); }
});

test('headless missing CJK substitutes locally while the same page uses the CDN', async () => {
  const requests: string[] = [];
  serveAssets(requests);
  setWorkerFontAssets(async (pack) => ({ urls: pack === 'latin' ? Object.fromEntries(latin.faces.map(({ file }) => [file, asset('latin', file).href])) : undefined }));
  const latinBytes = await worker.loadBundledFontBytes(fonts.resolveMetricCompatFace('Arial', false, false)!, undefined, runtime);
  requests.length = 0;
  await expect(worker.loadBundledFontBytes(japanese, undefined, { decoder, retain: false })).resolves.toEqual(latinBytes);
  expect(requests.every((url) => url.startsWith('file:'))).toBe(true);
  const restore = withBrowserFontGlobals();
  try { expect(await fonts.bundledFaceInstalled(japanese)).toBe(true); }
  finally { restore(); }
});

test('a responsive worker inherits the page default and its later opt-out', async () => {
  const restore = withBrowserFontGlobals();
  const channel = createWorkerFontPacks((message) => { page.receive(message); });
  const page = createPageFontPackResponder(channel.receive);
  setWorkerFontAssets(channel.source);
  const requests: string[] = [];
  serveAssets(requests);
  try {
    await expect(worker.loadBundledFontBytes(regular, undefined, runtime)).resolves.toBeInstanceOf(ArrayBuffer);
    expect(requests).toEqual([fontCdnAssetUrl(latin, regular.file).href]);
    fonts.configureFontCdn({ baseUrl: false });
    await expect(worker.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('install @betteroffice/fonts-latin');
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('install @betteroffice/fonts-latin');
    expect(requests).toHaveLength(1);
  } finally { channel.dispose(); restore(); }
});

test('a silent page leaves its worker with exactly the CDN sent at open', async () => {
  const restorePage = withBrowserFontGlobals();
  const fromOpen = getFontCdnConfiguration();
  restorePage();
  const restoreWorker = withBrowserFontGlobals(false);
  const clock = cdnClock();
  const posted = deferred<void>();
  const channel = createWorkerFontPacks(() => posted.resolve(), 5);
  setWorkerFontAssets(channel.source);
  const requests: string[] = [];
  serveAssets(requests);
  try {
    setFontCdnConfiguration(fromOpen);
    const load = worker.loadBundledFontBytes(regular, undefined, runtime);
    await beforeFontLoadFinishes(posted.promise, load, 'the worker pack request');
    await clock.advance(5);
    await expect(withinTestBudget(load, 'the silent-page font load')).resolves.toBeInstanceOf(ArrayBuffer);
    expect(requests).toEqual([fontCdnAssetUrl(latin, regular.file).href]);
    setFontCdnConfiguration(undefined);
    await expect(worker.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('install @betteroffice/fonts-latin');
    expect(requests).toHaveLength(1);
  } finally { channel.dispose(); clock.restore(); restoreWorker(); }
});

for (const failure of ['blocked', 403, 404, 410, 429, 500, 502, 503, 504, 'bad hash', 'bad length', 'corrupt compressed bytes'] as const) {
  test(`implicit ${failure} substitutes two faces, warns once and cools the failed scope`, async () => {
    const restore = withBrowserFontGlobals();
    const { warning, diagnostics, restore: stopObserving } = observeCdnFailures();
    const requests: string[] = [];
    const bold = fonts.resolveMetricCompatFace('Calibri', true, false)!;
    const base = fonts.createFontProvider(undefined, { decoder, retain: false }).resolveBaseLatin!;
    const decode: FontDecoder = async (encoded, options) => {
      const decoded = await decoder(encoded, options);
      if (failure === 'bad hash' && decoded.font.byteLength === regular.byteLength) new Uint8Array(decoded.font)[0] ^= 1;
      return decoded;
    };
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      if (!String(input).startsWith(jsDelivr)) return fontAssetResponse(input);
      requests.push(String(input));
      if (failure === 'blocked') throw new TypeError('Failed to fetch');
      if (typeof failure === 'number') return new Response(null, { status: failure });
      if (failure === 'bad length') return new Response(new Uint8Array(4));
      if (failure === 'corrupt compressed bytes') {
        const bytes = await (await fontAssetResponse(input)).arrayBuffer();
        return new Response(new Uint8Array(bytes.byteLength));
      }
      return fontAssetResponse(input);
    }) as typeof fetch;
    try {
      const mode = { decoder: decode, retain: false };
      await expect(fonts.loadBundledFontBytes(regular, undefined, mode)).resolves.toEqual(await base(false, false)!());
      await expect(fonts.loadBundledFontBytes(bold, undefined, mode)).resolves.toEqual(await base(true, false)!());
      await expect(fonts.loadBundledFontBytes(regular, undefined, { ...mode, strict: true })).rejects.toThrow();
      expect(requests).toEqual([regularUrl]);
      expect(warning.mock.calls).toHaveLength(1);
      for (const detail of ['cdn.jsdelivr.net', 'install @betteroffice/fonts-latin', 'configureFontCdn({ baseUrl: false })']) {
        expect(warning.mock.calls[0]?.[0]).toContain(detail);
      }
      expect(diagnostics).toEqual([
        expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' }),
        expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' }),
      ]);
      expect(await fonts.checkFontPacks()).toEqual([]);
    } finally {
      stopObserving();
      restore();
    }
  });
}

test.each([404, 'bad hash'] as const)('styled faces wait for a Regular dictionary with %s to recover', async (failure) => {
  const restore = withBrowserFontGlobals();
  const observed = observeCdnFailures();
  const now = Date.now();
  const clock = spyOn(Date, 'now').mockReturnValue(now);
  const bold = fonts.resolveMetricCompatFace('Calibri', true, false)!;
  const boldUrl = fontCdnAssetUrl(latin, bold.file).href;
  const requests: string[] = [];
  let broken = true;
  const decode: FontDecoder = async (encoded, options) => {
    const decoded = await decoder(encoded, options);
    if (broken && failure === 'bad hash') new Uint8Array(decoded.font)[0] ^= 1;
    return decoded;
  };
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    if (!url.startsWith(jsDelivr)) return fontAssetResponse(input);
    requests.push(url);
    return broken && failure === 404 && url === regularUrl
      ? new Response(null, { status: 404 }) : fontAssetResponse(input);
  }) as typeof fetch;
  const mode = { ...runtime, decoder: decode };
  const reason = failure === 404 ? 'HTTP 404' : 'SHA-256 mismatch';
  try {
    await expect(fonts.loadBundledFontBytes(regular, undefined, mode)).rejects.toThrow(reason);
    broken = false;
    for (const time of [now, now + cooldownMs - 1]) {
      clock.mockReturnValue(time);
      await expect(fonts.loadBundledFontBytes(bold, undefined, mode)).rejects.toThrow(reason);
      expect(requests).toEqual([regularUrl]);
    }
    expect(observed.warning).toHaveBeenCalledTimes(1);
    clock.mockReturnValue(now + cooldownMs);
    await expect(fonts.loadBundledFontBytes(bold, undefined, mode)).resolves.toBeInstanceOf(ArrayBuffer);
    expect(requests).toEqual([regularUrl, boldUrl, regularUrl]);
  } finally {
    clock.mockRestore();
    observed.restore();
    restore();
  }
});

test.each(['aborted', 'opaque', 'error', 'non-OK'] as const)('an implicit %s response cancels its unread body', async (failure) => {
  const observed = observeCdnFailures();
  const clock = cdnClock();
  const reply = deferred<Response>();
  const started = deferred<void>();
  let cancelled = false;
  const response = new Response(new ReadableStream({ cancel() { cancelled = true; } }), { status: 403 });
  if (failure === 'opaque' || failure === 'error') Object.defineProperty(response, 'type', { value: failure });
  globalThis.fetch = (async () => { started.resolve(); return reply.promise; }) as typeof fetch;
  try {
    const load = fontCdn.loadImplicitFontAsset('font', new URL(regularUrl), 'latin', async (bytes) => bytes);
    await beforeFontLoadFinishes(started.promise, load, 'the cancellable CDN request');
    if (failure === 'aborted') {
      await clock.advance(stallMs);
    }
    reply.resolve(response);
    await expect(withinTestBudget(load, 'the response failure')).rejects
      .toThrow(failure === 'aborted' ? 'No font data' : failure === 'non-OK' ? 'HTTP 403' : 'No readable response');
    await clock.advance(0);
    expect(cancelled).toBe(true);
  } finally {
    reply.resolve(response);
    clock.restore();
    observed.restore();
  }
});

test('concurrent faces and dictionaries share the network failure of their implicit base', async () => {
  const restore = withBrowserFontGlobals();
  const observed = observeCdnFailures();
  const bold = fonts.resolveMetricCompatFace('Calibri', true, false)!;
  const thai = fonts.resolveScriptFallbackFace('thai', false, false)!;
  const requests: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requests.push(String(input));
    throw new TypeError('Offline');
  }) as typeof fetch;
  try {
    const first = await Promise.allSettled([
      fonts.loadBundledFontBytes(regular, undefined, runtime),
      fonts.loadBundledFontBytes(bold, undefined, runtime),
    ]);
    expect(first.map(({ status }) => status)).toEqual(['rejected', 'rejected']);
    expect(requests).toHaveLength(1);
    await expect(fonts.loadBundledFontBytes(thai, undefined, runtime)).rejects.toThrow('Offline');
    await expect(fonts.loadBundledFontBytes(thai, undefined, runtime)).rejects.toThrow('Offline');
    expect(requests).toHaveLength(2);
    expect(requests[1]).toBe(fontCdnAssetUrl(intl, thai.file).href);
  } finally {
    observed.restore();
    restore();
  }
});

test('a valid implicit load preserves another URL’s 404 and warning suppression', async () => {
  const restore = withBrowserFontGlobals();
  const observed = observeCdnFailures();
  const other = fonts.resolveMetricCompatFace('Times New Roman', false, false)!;
  const otherUrl = fontCdnAssetUrl(latin, other.file).href;
  const now = Date.now();
  const clock = spyOn(Date, 'now').mockReturnValue(now);
  const requests: string[] = [];
  let available = false;
  let offline = false;
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requests.push(String(input));
    if (offline) throw new TypeError('Offline');
    return !available && String(input) === regularUrl
      ? new Response(null, { status: 404 }) : fontAssetResponse(input);
  }) as typeof fetch;
  try {
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('HTTP 404');
    available = true;
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('HTTP 404');
    expect(requests).toEqual([regularUrl]);
    await expect(fonts.loadBundledFontBytes(other, undefined, runtime)).resolves.toBeInstanceOf(ArrayBuffer);
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('HTTP 404');
    expect(requests).toEqual([regularUrl, otherUrl]);
    offline = true;
    await expect(fonts.loadBundledFontBytes(other, undefined, runtime)).rejects.toThrow('Offline');
    expect(requests).toEqual([regularUrl, otherUrl, otherUrl]);
    expect(observed.warning).toHaveBeenCalledTimes(1);

    clock.mockReturnValue(now + cooldownMs);
    offline = false;
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).resolves.toBeInstanceOf(ArrayBuffer);
    offline = true;
    await expect(fonts.loadBundledFontBytes(other, undefined, runtime)).rejects.toThrow('Offline');
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('Offline');
    expect(requests).toEqual([regularUrl, otherUrl, otherUrl, regularUrl, otherUrl]);
    expect(observed.warning).toHaveBeenCalledTimes(2);
  } finally {
    clock.mockRestore();
    observed.restore();
    restore();
  }
});

test('configureFontCdn clears the cooldown and installed fonts remain available while it is active', async () => {
  const restore = withBrowserFontGlobals();
  const observed = observeCdnFailures();
  const requests: string[] = [];
  let availableLocally = false;
  const { loadBundledFontBytes } = createFontRuntime(async (pack) => ({
    cdn: getFontCdnConfiguration(),
    urls: pack === 'latin' && availableLocally ? { [regular.file]: () => asset('latin', regular.file) } : undefined,
  }));
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    if (!String(input).startsWith(jsDelivr)) return fontAssetResponse(input);
    requests.push(String(input));
    throw new TypeError('Offline');
  }) as typeof fetch;
  try {
    await expect(loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('Offline');
    availableLocally = true;
    expect(hash(await loadBundledFontBytes(regular, undefined, runtime))).toBe(regularHash);
    availableLocally = false;
    await expect(loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('Offline');
    expect(requests).toEqual([regularUrl]);
    expect(observed.warning).toHaveBeenCalledTimes(1);

    fonts.configureFontCdn({ baseUrl: false });
    setFontCdnConfiguration(undefined);
    await expect(loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('Offline');
    expect(requests).toEqual([regularUrl, regularUrl]);
    expect(observed.warning).toHaveBeenCalledTimes(2);
  } finally {
    availableLocally = false;
    observed.restore();
    restore();
  }
});

test('independently loaded font CDN modules share the realm cooldown', async () => {
  const page: typeof fontCdn = await import(new URL('./fontCdn.ts?page', import.meta.url).href);
  const worker: typeof fontCdn = await import(new URL('./fontCdn.ts?worker', import.meta.url).href);
  expect(page.loadImplicitFontAsset).not.toBe(worker.loadImplicitFontAsset);

  const observed = observeCdnFailures();
  const fetchFont = mock(async () => { throw new TypeError('Offline'); });
  globalThis.fetch = fetchFont as unknown as typeof fetch;
  const url = new URL(regularUrl);
  const load = async (bytes: ArrayBuffer) => bytes;
  try {
    await expect(page.loadImplicitFontAsset(regular.file, url, 'latin', load)).rejects.toThrow('Offline');
    await expect(worker.loadImplicitFontAsset(regular.file, url, 'latin', load)).rejects.toThrow('Offline');
    expect(fetchFont).toHaveBeenCalledTimes(1);
    expect(observed.warning).toHaveBeenCalledTimes(1);
  } finally {
    observed.restore();
  }
});

test('a worker learns base failures independently across whole faces and CJK slices', async () => {
  const thread = new Worker(new URL('./cdnWorker.test-support.ts', import.meta.url));
  const restore = withBrowserFontGlobals();
  const observed = observeCdnFailures();
  const requests: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requests.push(String(input));
    throw new TypeError('Offline');
  }) as typeof fetch;
  try {
    const response = new Promise<{ requests: string[]; statuses: string[]; cjkRequests: string[]; warnings: number }>((resolve, reject) => {
      thread.once('error', reject);
      thread.once('exit', (code) => reject(new Error(`Font worker exited before replying (${code})`)));
      thread.once('message', resolve);
      thread.postMessage(getFontCdnConfiguration());
    });
    const [result] = await Promise.all([
      withinTestBudget(response, 'the font worker reply', 20_000),
      expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('Offline'),
    ]);
    expect(result.statuses).toEqual(['rejected', 'rejected', 'rejected']);
    expect(result.requests).toEqual([regularUrl]);
    expect(result.cjkRequests).toHaveLength(1);
    expect(result.cjkRequests[0]).toMatch(/\/c\d+\.zst$/);
    expect(result.warnings).toBe(2);
    expect(requests).toEqual([regularUrl]);
  } finally {
    observed.restore();
    restore();
    thread.unref();
    await withinTestBudget(thread.terminate(), 'the font worker to terminate');
  }
}, 30_000);

test('the published configureFontCdn clears the progressive entry cooldown', async () => {
  const published = await import('@betteroffice/fonts');
  const progressive = await import('@betteroffice/fonts/progressive');
  const previous = getFontCdnConfiguration();
  published.configureFontCdn({ baseUrl: false });
  const observed = observeCdnFailures();
  const requests: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requests.push(String(input));
    throw new TypeError('Offline');
  }) as typeof fetch;
  try {
    const url = Object.assign(async (file: string) => fontCdnAssetUrl(cjk, file), { implicitCdn: () => true });
    const slices = {
      ...fonts.createFontProvider().resolveScriptFallback('cjk-jp', false, false)!.slices!,
      prepareUrls: undefined, base: jsDelivr, url,
    };
    await expect(progressive.openSlicedFace(slices)).rejects.toThrow('Offline');
    await expect(progressive.openSlicedFace(slices)).rejects.toThrow('Offline');
    const index = fontCdnAssetUrl(cjk, slices.ref.index).href;
    expect(requests).toEqual([index]);

    published.configureFontCdn({ baseUrl: false });
    await expect(progressive.openSlicedFace(slices)).rejects.toThrow('Offline');
    expect(requests).toEqual([index, index]);
    expect(observed.warning).toHaveBeenCalledTimes(2);
  } finally {
    observed.restore();
    published.configureFontCdn({ baseUrl: false });
    setFontCdnConfiguration(previous);
  }
});

test.each(['index', 'skeleton', 'core', 'chunk'] as const)('an implicit CJK %s failure keeps its cooldown across successful requests', async (part) => {
  const restore = withBrowserFontGlobals();
  const observed = observeCdnFailures();
  const now = Date.now();
  const clock = spyOn(Date, 'now').mockReturnValue(now);
  const provider = fonts.createFontProvider(undefined, runtime);
  const slices = provider.resolveScriptFallback('cjk-jp', false, false)!.slices!;
  const url = (file: string) => fontCdnAssetUrl(cjk, file);
  const requests: string[] = [];
  let broken = true;
  try {
    const indexUrl = url(slices.ref.index);
    const { skeleton, core, chunks } = await (await fontAssetResponse(indexUrl)).json() as {
      skeleton: { file: string; encodedLength: number }; core: { file: string }; chunks: { file: string }[];
    };
    const files = [slices.ref.index, skeleton.file, core.file, chunks[0]!.file];
    const prefix = files.slice(0, ['index', 'skeleton', 'core', 'chunk'].indexOf(part) + 1).map((file) => url(file).href);
    const failedUrl = prefix.at(-1)!;
    const other = url('manifest.json');
    const missing = url('missing.zst');
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      const href = String(input);
      if (href.startsWith(jsDelivr)) requests.push(href);
      if (href === missing.href) return new Response(null, { status: 404 });
      if (broken && href === failedUrl) return part === 'skeleton'
        ? new Response(new Uint8Array(skeleton.encodedLength)) : new Response(null, { status: 404 });
      return fontAssetResponse(input);
    }) as typeof fetch;
    let load: () => Promise<unknown> = () => openSlicedFace(slices);
    if (part === 'chunk') {
      const face = await openSlicedFace(slices);
      const encoded = await (await fontAssetResponse(url(skeleton.file))).arrayBuffer();
      const parsed = parseSkeleton(new Uint8Array((await decoder(encoded, { codec: 'zstd' })).font));
      let glyph = 0;
      while (glyph < parsed.numGlyphs && parsed.chunkOfGlyph(glyph) !== 1) glyph++;
      expect(glyph).toBeLessThan(parsed.numGlyphs);
      load = () => face.ensureGlyphs([glyph]);
    }
    const error = { diagnostic: { pack: 'cjk', reason: 'cdn-unavailable' } };
    await expect(load()).rejects.toMatchObject(error);
    expect(requests).toEqual(prefix);
    broken = false;
    await fontCdn.loadImplicitFontAsset('manifest.json', other, 'cjk', async (bytes, check) =>
      check(() => JSON.parse(new TextDecoder().decode(bytes))));
    await expect(fontCdn.loadImplicitFontAsset('missing.zst', missing, 'cjk', async (bytes) => bytes)).rejects.toThrow('HTTP 404');
    clock.mockReturnValue(now + cooldownMs - 1);
    await expect(load()).rejects.toMatchObject(error);
    await expect(slices.loadWholeForExport!()).rejects.toMatchObject(error);
    expect(requests).toEqual([...prefix, other.href, missing.href]);
    expect(observed.warning).toHaveBeenCalledTimes(1);
    clock.mockReturnValue(now + cooldownMs);
    await expect(load()).resolves.toBeDefined();
    expect(requests).toEqual([...prefix, other.href, missing.href, ...(part === 'chunk' ? [failedUrl] : files.slice(0, 3).map((file) => url(file).href))]);
  } finally {
    clock.mockRestore();
    observed.restore();
    restore();
  }
}, 30_000);

test('an implicit CJK slice failure allows a later whole-face request and warns once', async () => {
  const restore = withBrowserFontGlobals();
  const { warning, diagnostics, restore: stopObserving } = observeCdnFailures();
  const requests: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    if (String(input).includes('/fonts-cjk@')) {
      requests.push(String(input));
      return new Response(null, { status: 404 });
    }
    return fontAssetResponse(input);
  }) as typeof fetch;
  try {
    const provider = fonts.createFontProvider(undefined, { decoder, retain: false });
    const load = provider.resolveScriptFallback('cjk-jp', false, false)!;
    await expect(openSlicedFace(load.slices!)).rejects.toThrow('HTTP 404');
    await expect(load()).resolves.toEqual(await provider.resolveBaseLatin!(false, false)!());
    expect(requests).toEqual([fontCdnAssetUrl(cjk, load.slices!.ref.index).href, fontCdnAssetUrl(cjk, japanese.file).href]);
    expect(warning.mock.calls).toHaveLength(1);
    expect(diagnostics).toEqual([
      expect.objectContaining({ pack: 'cjk', reason: 'cdn-unavailable' }),
      expect.objectContaining({ pack: 'cjk', reason: 'cdn-unavailable' }),
    ]);
  } finally {
    stopObserving();
    restore();
  }
});

test('T18: implicit CDN completes a stream that keeps making progress past the stall budget', async () => {
  const encoded = new Uint8Array(await Bun.file(asset('latin', regular.file)).arrayBuffer());
  const restore = withBrowserFontGlobals();
  const failure = observeCdnFailures();
  const clock = cdnClock();
  const pulls = [deferred<void>(), deferred<void>()] as const;
  globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
    if (String(input) !== regularUrl) return fontAssetResponse(input);
    return streamed(encoded, AbortSignal.any([options!.signal!, clock.signal]), pulls);
  }) as typeof fetch;
  try {
    const result = fonts.loadBundledFontBytes(regular, undefined, { decoder, retain: false });
    await beforeFontLoadFinishes(pulls[0].promise, result, 'the first stream read');
    await clock.advance(6_000);
    await beforeFontLoadFinishes(pulls[1].promise, result, 'the second stream read');
    await clock.advance(6_000);
    expect(hash(await withinTestBudget(result, 'the streamed face'))).toBe(regularHash);
    expect(failure.warning).not.toHaveBeenCalled();
    expect(failure.diagnostics).toEqual([]);
  } finally {
    clock.restore();
    failure.restore();
    restore();
  }
});

test('T19: a stalled body cools the base until the retry window opens', async () => {
  const restore = withBrowserFontGlobals();
  const failure = observeCdnFailures();
  const clock = cdnClock();
  const started = deferred<AbortSignal>();
  const pulling = deferred<void>();
  let requests = 0;
  globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
    if (String(input) !== regularUrl || ++requests > 1) return fontAssetResponse(input);
    const signal = options!.signal!;
    return new Response(new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array([1]));
        started.resolve(signal);
      },
      pull() {
        pulling.resolve();
        return waitForAbort(AbortSignal.any([signal, clock.signal]));
      },
    }));
  }) as typeof fetch;
  try {
    const mode = { decoder: ((bytes, options) => decoder(bytes, options)) satisfies FontDecoder };
    const base = fonts.createFontProvider(undefined, mode).resolveBaseLatin!(false, false)!;
    const fallback = await base();
    const result = fonts.loadBundledFontBytes(regular, undefined, mode);
    const signal = await beforeFontLoadFinishes(started.promise, result, 'the stalled stream');
    await beforeFontLoadFinishes(pulling.promise, result, 'the stalled body read');
    await clock.advance(0);
    await clock.advance(stallMs - 1);
    expect(signal.aborted).toBe(false);
    await clock.advance(1);
    expect(signal.aborted).toBe(true);
    await expect(withinTestBudget(result, 'the stalled body fallback')).resolves.toEqual(fallback);
    await clock.advance(cooldownMs - 1);
    await expect(withinTestBudget(fonts.loadBundledFontBytes(regular, undefined, mode), 'the cooldown fallback')).resolves.toEqual(fallback);
    expect(requests).toBe(1);
    await clock.advance(1);
    expect(hash(await fonts.loadBundledFontBytes(regular, undefined, mode))).toBe(regularHash);
    expect(requests).toBe(2);
    expect(failure.warning).toHaveBeenCalledTimes(1);
    expect(failure.diagnostics).toEqual([
      expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' }),
      expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' }),
    ]);
  } finally {
    clock.restore();
    failure.restore();
    restore();
  }
});

test('T20: the implicit stall budget starts before response headers', async () => {
  const restore = withBrowserFontGlobals();
  const failure = observeCdnFailures();
  const clock = cdnClock();
  const started = deferred<AbortSignal>();
  globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
    if (String(input) !== regularUrl) return fontAssetResponse(input);
    const signal = options!.signal!;
    started.resolve(signal);
    return waitForAbort(AbortSignal.any([signal, clock.signal]));
  }) as typeof fetch;
  try {
    const base = fonts.createFontProvider(undefined, runtime).resolveBaseLatin!(false, false)!;
    const fallback = await base();
    const result = fonts.loadBundledFontBytes(regular, undefined, { decoder, retain: false });
    const signal = await beforeFontLoadFinishes(started.promise, result, 'the request without headers');
    await clock.advance(stallMs - 1);
    expect(signal.aborted).toBe(false);
    await clock.advance(1);
    expect(signal.aborted).toBe(true);
    await expect(withinTestBudget(result, 'the stalled headers fallback')).resolves.toEqual(fallback);
    expect(failure.warning).toHaveBeenCalledTimes(1);
    expect(failure.diagnostics).toEqual([expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' })]);
  } finally {
    clock.restore();
    failure.restore();
    restore();
  }
});

test('T21: an implicit response without a readable body gets one bounded body wait', async () => {
  const encoded = await Bun.file(asset('latin', regular.file)).arrayBuffer();
  const restore = withBrowserFontGlobals();
  const failure = observeCdnFailures();
  const clock = cdnClock();
  const started = deferred<AbortSignal>();
  const headers = deferred<Response>();
  const body = deferred<ArrayBuffer>();
  const reading = deferred<void>();
  const arrayBuffer = mock(() => { reading.resolve(); return body.promise; });
  const response = { ok: true, body: null, arrayBuffer } as unknown as Response;
  globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
    if (String(input) !== regularUrl) return fontAssetResponse(input);
    started.resolve(options!.signal!);
    return headers.promise;
  }) as typeof fetch;
  try {
    const base = fonts.createFontProvider(undefined, runtime).resolveBaseLatin!(false, false)!;
    const fallback = await base();
    const result = fonts.loadBundledFontBytes(regular, undefined, { decoder, retain: false });
    const signal = await beforeFontLoadFinishes(started.promise, result, 'the request without a readable body');
    await clock.advance(2_000);
    headers.resolve(response);
    await beforeFontLoadFinishes(reading.promise, result, 'the arrayBuffer call');
    expect(arrayBuffer).toHaveBeenCalledTimes(1);
    await clock.advance(stallMs - 1);
    expect(signal.aborted).toBe(false);
    await clock.advance(1);
    expect(signal.aborted).toBe(true);
    body.resolve(encoded);
    await expect(withinTestBudget(result, 'the unreadable body fallback')).resolves.toEqual(fallback);
    expect(failure.warning).toHaveBeenCalledTimes(1);
    expect(failure.diagnostics).toEqual([expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' })]);
  } finally {
    headers.resolve(response);
    body.resolve(encoded);
    clock.restore();
    failure.restore();
    restore();
  }
});

test('T22: only the explicitly configured CDN applies a total timeout to a progressing stream', async () => {
  const encoded = new Uint8Array(await Bun.file(asset('latin', regular.file)).arrayBuffer());
  const restore = withBrowserFontGlobals();
  const failure = observeCdnFailures();
  const clock = cdnClock();
  const explicitStarted = deferred<AbortSignal>();
  const implicitStarted = deferred<AbortSignal>();
  const explicitPulls = [deferred<void>(), deferred<void>()] as const;
  const implicitPulls = [deferred<void>(), deferred<void>()] as const;
  globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
    const url = String(input);
    if (!url.startsWith('https://explicit.example/') && url !== regularUrl) return fontAssetResponse(input);
    const signal = options!.signal!;
    (url === regularUrl ? implicitStarted : explicitStarted).resolve(signal);
    return streamed(encoded, AbortSignal.any([signal, clock.signal]), url === regularUrl ? implicitPulls : explicitPulls);
  }) as typeof fetch;
  try {
    const timeoutMs = 9_500;
    fonts.configureFontCdn({ baseUrl: 'https://explicit.example/', timeoutMs });
    const explicit = fonts.loadBundledFontBytes(regular, undefined, runtime).catch((error: unknown) => error);
    const explicitSignal = await beforeFontLoadFinishes(explicitStarted.promise, explicit, 'the explicit CDN request');
    await beforeFontLoadFinishes(explicitPulls[0].promise, explicit, 'the first explicit stream read');
    await clock.advance(6_000);
    await beforeFontLoadFinishes(explicitPulls[1].promise, explicit, 'the second explicit stream read');
    expect(explicitSignal.aborted).toBe(false);
    await clock.advance(timeoutMs - 6_000);
    expect(explicitSignal.aborted).toBe(true);
    expect(await withinTestBudget(explicit, 'the explicit CDN timeout')).toMatchObject({ name: 'TimeoutError' });

    setFontCdnConfiguration(undefined);
    const implicit = fonts.loadBundledFontBytes(regular, undefined, { decoder, retain: false });
    const implicitSignal = await beforeFontLoadFinishes(implicitStarted.promise, implicit, 'the implicit CDN request');
    await beforeFontLoadFinishes(implicitPulls[0].promise, implicit, 'the first implicit stream read');
    await clock.advance(6_000);
    await beforeFontLoadFinishes(implicitPulls[1].promise, implicit, 'the second implicit stream read');
    expect(implicitSignal.aborted).toBe(false);
    await clock.advance(timeoutMs - 6_000);
    expect(implicitSignal.aborted).toBe(false);
    await clock.advance(12_000 - timeoutMs);
    expect(hash(await withinTestBudget(implicit, 'the implicit CDN stream'))).toBe(regularHash);
    expect(failure.warning).not.toHaveBeenCalled();
    expect(failure.diagnostics).toEqual([]);
  } finally {
    clock.restore();
    failure.restore();
    restore();
  }
});

test('an explicit CDN remains retryable after the implicit CDN has failed', async () => {
  const restore = withBrowserFontGlobals();
  const { warning, restore: stopObserving } = observeCdnFailures();
  const failure = new TypeError('Blocked font request');
  const requests: string[] = [];
  let explicitWorks = false;
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    if (!url.startsWith('https:')) return fontAssetResponse(input);
    requests.push(url);
    if (explicitWorks && url.startsWith('https://explicit.example/')) return fontAssetResponse(input);
    throw failure;
  }) as typeof fetch;
  try {
    await expect(fonts.loadBundledFontBytes(regular, undefined, { decoder, retain: false })).resolves.toBeInstanceOf(ArrayBuffer);
    fonts.configureFontCdn({ baseUrl: 'https://explicit.example/' });
    await expect(fonts.loadBundledFontBytes(regular, undefined, { decoder, retain: false })).rejects.toBe(failure);
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toBe(failure);
    explicitWorks = true;
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).resolves.toBeInstanceOf(ArrayBuffer);
    const explicit = `https://explicit.example/latin/${latin.version}/${regular.file}`;
    expect(requests).toEqual([fontCdnAssetUrl(latin, regular.file).href, explicit, explicit, explicit]);
    expect(warning.mock.calls).toHaveLength(1);
  } finally {
    stopObserving();
    restore();
  }
});

test('a page CJK slice loader uses the CDN configured after its creation', async () => {
  fonts.configureFontCdn({ baseUrl: 'https://old.example/' });
  const slices = fonts.createFontProvider(undefined, { decoder, retain: false }).resolveScriptFallback('cjk-jp', false, false)!.slices!;
  fonts.configureFontCdn({ baseUrl: 'https://new.example/' });
  const requests: string[] = [];
  serveAssets(requests);
  await openSlicedFace(slices);
  expect(requests).toContain(`https://new.example/cjk/${cjk.version}/${slices.ref.index}`);
  expect(requests.some((url) => url.startsWith('https://old.example/'))).toBe(false);
});

test('concurrent asset lookups for one worker pack coalesce on the wire', async () => {
  const clock = cdnClock();
  const posts = mock(() => {});
  const channel = createWorkerFontPacks(posts, 5);
  try {
    const lookups = Promise.all([channel.source('cjk'), channel.source('cjk')]);
    await clock.advance(5);
    await withinTestBudget(lookups, 'the coalesced worker pack lookups');
    expect(posts).toHaveBeenCalledTimes(1);
  } finally { channel.dispose(); clock.restore(); }
});
test('installed packs beat the browser default without a CDN request', async () => {
  const restore = withBrowserFontGlobals();
  const requests: string[] = [];
  serveAssets(requests);
  installed = true;
  try {
    const bytes = await fonts.loadBundledFontBytes(regular, undefined, runtime);
    expect(hash(bytes)).toBe(latin.faces.find(({ file }) => file === regular.file)!.sha256);
    expect(requests.every((url) => url.startsWith('file:'))).toBe(true);
    expect(getFontCdnConfiguration()).toMatchObject(implicitCdn);
  } finally { installed = false; restore(); }
});

test('installed packs win over a configured CDN without changing their request policy', async () => {
  installed = true;
  fonts.configureFontCdn({ baseUrl: 'https://unused.example/' });
  const requests: string[] = [];
  serveAssets(requests, (_url, options) => expect(options).toBeUndefined());
  for (const [kind, file] of [
    ['latin', regular.file], ['intl', fonts.resolveScriptFallbackFace('thai', false, false)!.file], ['cjk', japanese.file],
  ] as const) expect((await fonts.fontAssetUrls(kind))?.[file]).toBe(asset(kind, file).href);
  for (const face of [regular, fonts.resolveScriptFallbackFace('thai', false, false)!, japanese]) {
    expect((await fonts.loadBundledFontBytes(face, undefined, runtime)).byteLength).toBe(face.byteLength);
  }
  const location = fonts.createFontProvider(undefined, runtime).resolveScriptFallback('cjk-jp', false, false)!.slices!;
  const sliced = await openSlicedFace(location);
  await sliced.ensureText('日本語');
  expect((await sliceSourceOf(sliced)!).timeoutMs).toBeUndefined();
  expect(requests.every((url) => !url.startsWith('https://unused.example/'))).toBe(true);
});

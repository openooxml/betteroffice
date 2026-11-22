import { afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import latin from '../../fonts-latin/assets/manifest.json';
import intl from '../../fonts-intl/assets/manifest.json';
import cjk from '../../fonts-cjk/assets/manifest.json';
import { decodeFontFace, type FontDecoder } from './decode';
import * as fontCdn from './fontCdn';
import { createPageFontPackResponder } from '../../../shared/office-session/fontPackSource';
import { fontAssetResponse, withBrowserFontGlobals } from './test-assets';
import { openSlicedFace, sliceSourceOf } from './progressive-entry';
import { parseSkeleton } from './slices';
import { writeFontCdnTree } from '../../../scripts/fonts/cdn-tree';
import { getFontCdnConfiguration, setFontCdnConfiguration, type FontCdnConfiguration } from '../../../shared/office-session/fontCdn';
import { createFontPackResponder, createWorkerFontPacks } from '../../../shared/office-session/optionalFontPacks';
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
});

afterEach(() => {
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
    let calls = 0;
    globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
      if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
      calls++;
      if (calls === 1) {
        if (failure === '404') return new Response(null, { status: 404 });
        return new Promise<Response>((_resolve, reject) => {
          const signal = options!.signal!;
          signal.addEventListener('abort', () => reject(signal.reason), { once: true });
        });
      }
      return new Response(await Bun.file(asset('latin', regular.file)).arrayBuffer());
    }) as typeof fetch;
    await expect(fonts.loadBundledFontBytes(regular, undefined, { decoder, strict: true })).rejects.toThrow();
    expect((await fonts.loadBundledFontBytes(regular, undefined, { decoder, strict: true })).byteLength).toBe(regular.byteLength);
    expect(calls).toBe(2);
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
const implicitCdn = { baseUrl: jsDelivr, timeoutMs: 3000, implicit: true } as const;
const cdnUnavailable = Symbol.for('betteroffice.fonts.cdn-unavailable');

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
  const channel = createWorkerFontPacks(() => {}, 5);
  setWorkerFontAssets(channel.source);
  const requests: string[] = [];
  serveAssets(requests);
  try {
    setFontCdnConfiguration(fromOpen);
    await expect(worker.loadBundledFontBytes(regular, undefined, runtime)).resolves.toBeInstanceOf(ArrayBuffer);
    expect(requests).toEqual([fontCdnAssetUrl(latin, regular.file).href]);
    setFontCdnConfiguration(undefined);
    await expect(worker.loadBundledFontBytes(regular, undefined, runtime)).rejects.toThrow('install @betteroffice/fonts-latin');
    expect(requests).toHaveLength(1);
  } finally { channel.dispose(); restoreWorker(); }
});

for (const failure of ['blocked', 'timeout', 'HTTP 404', 'bad hash', 'bad length'] as const) {
  test(`implicit ${failure} substitutes two faces, warns once and stops requests for the pack`, async () => {
    Reflect.deleteProperty(globalThis, cdnUnavailable);
    const restore = withBrowserFontGlobals();
    const warning = spyOn(console, 'warn').mockImplementation(() => {});
    const timeout = AbortSignal.timeout.bind(AbortSignal);
    const shortened = spyOn(AbortSignal, 'timeout').mockImplementation((ms) => timeout(ms === 3000 ? 5 : ms));
    const requests: string[] = [];
    const bold = fonts.resolveMetricCompatFace('Calibri', true, false)!;
    const base = fonts.createFontProvider(undefined, { decoder, retain: false }).resolveBaseLatin!;
    const decode: FontDecoder = async (encoded, options) => {
      const decoded = await decoder(encoded, options);
      if (failure === 'bad hash' && decoded.font.byteLength === regular.byteLength) new Uint8Array(decoded.font)[0] ^= 1;
      return decoded;
    };
    globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
      if (!String(input).startsWith(jsDelivr)) return fontAssetResponse(input);
      requests.push(String(input));
      if (failure === 'blocked') throw new TypeError('Failed to fetch');
      if (failure === 'HTTP 404') return new Response(null, { status: 404 });
      if (failure === 'bad length') return new Response(new Uint8Array(4));
      if (failure === 'timeout') {
        const signal = options!.signal!;
        return new Promise<Response>((_resolve, reject) => {
          if (signal.aborted) reject(signal.reason);
          else signal.addEventListener('abort', () => reject(signal.reason), { once: true });
        });
      }
      return fontAssetResponse(input);
    }) as typeof fetch;
    try {
      const mode = { decoder: decode, retain: false };
      await expect(fonts.loadBundledFontBytes(regular, undefined, mode)).resolves.toEqual(await base(false, false)!());
      await expect(fonts.loadBundledFontBytes(bold, undefined, mode)).resolves.toEqual(await base(true, false)!());
      await expect(fonts.loadBundledFontBytes(regular, undefined, { ...mode, strict: true })).rejects.toThrow();
      expect(requests).toEqual([fontCdnAssetUrl(latin, regular.file).href]);
      expect(shortened).toHaveBeenCalledWith(3000);
      expect(warning.mock.calls).toHaveLength(1);
      for (const detail of ['cdn.jsdelivr.net', 'install @betteroffice/fonts-latin', 'configureFontCdn({ baseUrl: false })']) {
        expect(warning.mock.calls[0]?.[0]).toContain(detail);
      }
      expect(await fonts.checkFontPacks()).toEqual([expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' })]);
    } finally {
      Reflect.deleteProperty(globalThis, cdnUnavailable);
      warning.mockRestore();
      shortened.mockRestore();
      restore();
    }
  });
}

test('an implicit CJK slice failure substitutes the whole face without another CDN request', async () => {
  Reflect.deleteProperty(globalThis, cdnUnavailable);
  const restore = withBrowserFontGlobals();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
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
    expect(requests).toEqual([fontCdnAssetUrl(cjk, load.slices!.ref.index).href]);
    expect(warning.mock.calls).toHaveLength(1);
    expect(await fonts.checkFontPacks()).toEqual([expect.objectContaining({ pack: 'cjk', reason: 'cdn-unavailable' })]);
  } finally {
    Reflect.deleteProperty(globalThis, cdnUnavailable);
    warning.mockRestore();
    restore();
  }
});

test('an explicit CDN remains retryable after the implicit CDN has failed', async () => {
  Reflect.deleteProperty(globalThis, cdnUnavailable);
  const restore = withBrowserFontGlobals();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
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
    explicitWorks = true;
    await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).resolves.toBeInstanceOf(ArrayBuffer);
    const explicit = `https://explicit.example/latin/${latin.version}/${regular.file}`;
    expect(requests).toEqual([fontCdnAssetUrl(latin, regular.file).href, explicit, explicit]);
    expect(warning.mock.calls).toHaveLength(1);
  } finally {
    Reflect.deleteProperty(globalThis, cdnUnavailable);
    warning.mockRestore();
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
  const posts = mock(() => {});
  const channel = createWorkerFontPacks(posts, 5);
  try {
    await Promise.all([channel.source('cjk'), channel.source('cjk')]);
    expect(posts).toHaveBeenCalledTimes(1);
  } finally { channel.dispose(); }
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
    expect(getFontCdnConfiguration()).toEqual(implicitCdn);
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

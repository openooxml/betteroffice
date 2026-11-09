import { afterEach, expect, mock, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import latin from '../../fonts-latin/assets/manifest.json';
import intl from '../../fonts-intl/assets/manifest.json';
import cjk from '../../fonts-cjk/assets/manifest.json';
import { decodeFontFace, type FontDecoder } from './decode';
import { fontCdnAssetUrl } from './fontCdn';
import { fontAssetResponse } from './test-assets';
import { openSlicedFace, sliceSourceOf } from './progressive-entry';
import { parseSkeleton } from './slices';
import { writeFontCdnTree } from '../../../scripts/fonts/cdn-tree';

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
    const match = url.pathname.match(/\/(latin|intl|cjk)\/[^/]+\/(.+)$/);
    if (!match) return fontAssetResponse(input);
    return new Response(await Bun.file(asset(match[1]!, match[2]!)).arrayBuffer());
  }) as typeof fetch;
}

afterEach(() => {
  fonts.configureFontCdn({ baseUrl: false });
  globalThis.fetch = originalFetch;
});

test('CDN loading starts disabled and preserves missing-pack errors', async () => {
  expect(fonts.fontCdnConfiguration()).toEqual({ baseUrl: false, timeoutMs: 30_000 });
  expect(await fonts.bundledFaceInstalled(regular)).toBe(false);
  await expect(fonts.loadBundledFontBytes(regular, undefined, runtime)).rejects.toMatchObject({ name: 'MissingFontAddonError' });
  await expect(fonts.loadBundledFontBytes(japanese, undefined, runtime)).rejects.toThrow('install @betteroffice/fonts-cjk');
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

test('installed packs win over a configured CDN without changing their request policy', async () => {
  installed = true;
  fonts.configureFontCdn({ baseUrl: 'https://unused.example/' });
  const requests: string[] = [];
  serveAssets(requests, (_url, options) => expect(options).toBeUndefined());
  for (const face of [regular, fonts.resolveScriptFallbackFace('thai', false, false)!, japanese]) {
    expect((await fonts.loadBundledFontBytes(face, undefined, runtime)).byteLength).toBe(face.byteLength);
  }
  const location = fonts.createFontProvider(undefined, runtime).resolveScriptFallback('cjk-jp', false, false)!.slices!;
  const sliced = await openSlicedFace(location);
  await sliced.ensureText('日本語');
  expect((await sliceSourceOf(sliced)!).timeoutMs).toBeUndefined();
  expect(requests.every((url) => !url.startsWith('https://unused.example/'))).toBe(true);
});

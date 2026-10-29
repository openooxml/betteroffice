import { afterEach, expect, mock, test } from 'bun:test';
import { createHash } from 'node:crypto';
import pack from '../../fonts-cjk/assets/manifest.json';
import cjkPackage from '../../fonts-cjk/package.json';
import { CJK_FONT_ASSET_URLS } from '../../fonts-cjk/src';
import { createFontProvider as cdn } from './cdn';
import type { FontDecoder } from './decode';
import { BUNDLED_FONTS, createFontProvider, loadBundledFontBytes } from './index';
import { isCjkFont, loadPackedFontBytes } from './packed';
import { fontAssetResponse } from './test-assets';

const cjk = BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-'));
const resolveUrl = async (face: { file: string }) => CJK_FONT_ASSET_URLS[face.file]();
const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
});

test('ships exactly five literal CJK URLs and a schema 1 zstd pack with pinned source hashes', async () => {
  expect(cjkPackage.files).toEqual(['dist', 'assets', 'LICENSE', 'LICENSES', 'README.md']);
  expect(pack.format).toBe(1);
  expect(pack.schema).toBe(1);
  expect(pack.pack).toBe('cjk');
  expect(pack.version).toMatch(/^\d+\.\d+\.\d+$/);
  expect(cjk).toHaveLength(5);
  const files = cjk.map((face) => face.file).sort();
  expect(pack.faces.map((entry) => entry.file)).toEqual(files);
  expect(Object.keys(CJK_FONT_ASSET_URLS).sort()).toEqual(files);
  const source = await Bun.file(new URL('../../fonts-cjk/src/index.ts', import.meta.url)).text();
  const rows = [...source.matchAll(/'([^']+)': \(\) => new URL\('\.\.\/assets\/([^']+)', import\.meta\.url\)/g)];
  expect(rows.map(([, file]) => file).sort()).toEqual(files);
  for (const [, file, target] of rows) {
    expect(file).toBe(target);
    expect(CJK_FONT_ASSET_URLS[file]().pathname).toEndWith(`/assets/${file}`);
  }
  const checksums = await Bun.file(new URL('../../../scripts/fonts/cjk-sources.sha256', import.meta.url)).text();
  const pinned = checksums.trim().split('\n').map((line) => line.trim().split(/\s+/));
  expect(pinned.map(([, file]) => `${file}.zst`).sort()).toEqual(files);
  for (const entry of pack.faces) {
    expect(isCjkFont(entry.file)).toBe(true);
    expect(entry.codec).toBe('zstd');
    expect(entry.dict).toBeNull();
    expect(entry.source).toBe(entry.file.replace(/\.zst$/, ''));
    expect(entry.sha256).toBe(pinned.find(([, file]) => file === entry.source)![0]);
  }
});

test('decodes every real CJK face through the packed loader and optional add-on to its manifest hash', async () => {
  for (const face of cjk) {
    const entry = pack.faces.find((entry) => entry.file === face.file)!;
    const first = await loadPackedFontBytes(face, resolveUrl);
    const second = await loadBundledFontBytes(face);
    expect(face.byteLength).toBe(entry.byteLength);
    expect(first.byteLength).toBe(entry.byteLength);
    expect(createHash('sha256').update(new Uint8Array(first)).digest('hex')).toBe(entry.sha256);
    expect(new DataView(first).getUint32(0)).toBe(0x4f54544f);
    expect(first).toEqual(second);
    expect(first).not.toBe(second);
    structuredClone(first, { transfer: [first] });
    expect(await loadPackedFontBytes(face, resolveUrl)).toEqual(second);
  }
});

test('bundled and CDN CJK providers use the bound engine decoder without fetching fontpack.wasm', async () => {
  const face = cjk.find((face) => face.family === 'Noto Sans JP')!;
  const expected = await loadPackedFontBytes(face, resolveUrl);
  const encoded = await Bun.file(CJK_FONT_ASSET_URLS[face.file]()).arrayBuffer();
  const decoder = mock<FontDecoder>(async (input, options) => {
    expect(input instanceof Uint8Array ? input : new Uint8Array(input)).toEqual(new Uint8Array(encoded));
    expect(options).toEqual({ codec: 'zstd', dict: undefined, maxOutput: 64 * 1024 * 1024 });
    return { font: expected.slice(0) };
  });
  const requested: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    requested.push(url);
    if (url.endsWith('fontpack.wasm')) throw new Error('standalone decoder requested');
    return fontAssetResponse(input);
  }) as typeof fetch;
  const runtime = { decoder };
  for (const provider of [
    createFontProvider(undefined, runtime),
    createFontProvider({ baseUrl: 'https://cjk-bundled.example/fonts/' }, runtime),
    cdn(undefined, runtime),
    cdn({ cjkBaseUrl: 'https://cjk-cdn.example/fonts/' }, runtime),
  ]) {
    const load = provider.resolveFamily(face.family, false, false)!;
    const first = await load();
    expect(first).toEqual(expected);
    structuredClone(first, { transfer: [first] });
    expect(await load()).toEqual(expected);
  }
  expect(decoder).toHaveBeenCalledTimes(4);
  expect(requested.filter((url) => !url.startsWith('file:'))).toEqual([
    `https://cjk-bundled.example/fonts/${face.file}`,
    `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-cjk@${cjkPackage.version}/assets/${face.file}`,
    `https://cjk-cdn.example/fonts/${face.file}`,
  ]);
});

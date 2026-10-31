import { isLatinFont } from './packed';
import { fontAssetResponse } from './test-assets';
import { afterEach, expect, mock, test } from 'bun:test';

type LatinAddon = { LATIN_FONT_ASSET_URLS?: unknown };

const missingLatinAddon = (): LatinAddon => {
  throw new Error("Cannot find package '@betteroffice/fonts-latin'");
};
const importLatin = mock(missingLatinAddon);
mock.module('@betteroffice/fonts-latin', () => ({
  then(resolve: (addon: LatinAddon) => void, reject: (error: unknown) => void) {
    Promise.resolve().then(importLatin).then(resolve, reject);
  },
}));

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
  importLatin.mockReset();
  importLatin.mockImplementation(missingLatinAddon);
});

test('reports the missing Latin package for every Latin face and retries misses', async () => {
  const { BUNDLED_FONTS, loadBundledFontBytes } = await import('./index');
  let attempts = 0;
  for (const face of BUNDLED_FONTS.filter((candidate) => isLatinFont(candidate.file))) {
    for (const script of [face.script, undefined]) {
      await expect(loadBundledFontBytes({ ...face, script })).rejects.toThrow(
        `Bundled font ${face.file} needs @betteroffice/fonts-latin - install @betteroffice/fonts-latin`,
      );
      expect(importLatin).toHaveBeenCalledTimes(++attempts);
    }
  }
});

test('recognizes missing-module errors without hiding dependency or evaluation failures', async () => {
  const { loadBundledFontBytes, resolveMetricCompatFace } = await import('./index');
  const face = resolveMetricCompatFace('Arial', false, false)!;
  for (const message of [
    "Cannot find module '@betteroffice/fonts-latin'",
    'Failed to resolve module specifier "@betteroffice/fonts-latin"',
    'Error resolving module specifier "@betteroffice/fonts-latin"',
    'The specifier "@betteroffice/fonts-latin" could not be resolved',
    "Module name, '@betteroffice/fonts-latin' does not resolve",
  ]) {
    importLatin.mockImplementation(() => { throw new Error(message); });
    await expect(loadBundledFontBytes(face)).rejects.toThrow('install @betteroffice/fonts-latin');
  }
  for (const message of [
    "Cannot find package 'dependency-of-latin'",
    "Cannot find package '@betteroffice/fonts-latin-extra'",
    'Failed to fetch dynamically imported module',
    'latin module evaluation failed',
  ]) {
    const failure = new Error(message);
    importLatin.mockImplementation(() => { throw failure; });
    await expect(loadBundledFontBytes(face)).rejects.toBe(failure);
  }
});

test('rejects invalid Latin exports and retries', async () => {
  const { loadBundledFontBytes, resolveMetricCompatFace } = await import('./index');
  const face = resolveMetricCompatFace('Arial', false, false)!;
  let attempts = 0;
  for (const addon of [{}, { LATIN_FONT_ASSET_URLS: null }, { LATIN_FONT_ASSET_URLS: 'invalid' }]) {
    importLatin.mockImplementation(() => addon);
    for (let retry = 0; retry < 2; retry++) {
      await expect(loadBundledFontBytes(face)).rejects.toThrow(
        /^Invalid LATIN_FONT_ASSET_URLS export from @betteroffice\/fonts-latin$/,
      );
      expect(importLatin).toHaveBeenCalledTimes(++attempts);
    }
  }
});

test('custom bases bypass the package import and return owned buffers', async () => {
  const { loadBundledFontBytes, resolveMetricCompatFace } = await import('./index');
  const face = resolveMetricCompatFace('Calibri', false, false)!;
  const requested: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requested.push(String(input));
    return fontAssetResponse(input);
  }) as typeof fetch;
  const options = { baseUrl: 'https://without-latin-custom.example/' };
  const first = await loadBundledFontBytes(face, options);
  structuredClone(first, { transfer: [first] });
  expect((await loadBundledFontBytes(face, options)).byteLength).toBe(face.byteLength);
  expect(requested.filter((url) => !url.endsWith('.wasm'))).toEqual([`https://without-latin-custom.example/${face.file}`]);
  expect(importLatin).not.toHaveBeenCalled();
});

test('recovers after installing Latin and reports missing asset entries', async () => {
  const { loadBundledFontBytes, resolveMetricCompatFace } = await import('./index');
  const face = resolveMetricCompatFace('Cambria', false, false)!;
  await expect(loadBundledFontBytes(face)).rejects.toThrow('install @betteroffice/fonts-latin');
  importLatin.mockImplementation(() => ({
    LATIN_FONT_ASSET_URLS: {
      [face.file]: () => new URL('https://latin-installed.example/Caladea-Regular.ttf.zst'),
    },
  }));
  globalThis.fetch = fontAssetResponse as typeof fetch;
  expect((await loadBundledFontBytes(face)).byteLength).toBe(face.byteLength);
  const other = resolveMetricCompatFace('Arial', false, false)!;
  await expect(loadBundledFontBytes(other)).rejects.toThrow(
    `Unknown bundled font asset: ${other.file}`,
  );
});

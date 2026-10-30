import { isLatinFont } from './packed';
import { fontAssetResponse } from './test-assets';
import { afterEach, expect, mock, test } from 'bun:test';

type IntlAddon = { INTL_FONT_ASSET_URLS?: unknown };

const missingIntlAddon = (): IntlAddon => {
  throw new Error("Cannot find package '@betteroffice/fonts-intl'");
};
const importIntl = mock(missingIntlAddon);
mock.module('@betteroffice/fonts-intl', () => ({
  then(resolve: (addon: IntlAddon) => void, reject: (error: unknown) => void) {
    Promise.resolve().then(importIntl).then(resolve, reject);
  },
}));

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
  importIntl.mockReset();
  importIntl.mockImplementation(missingIntlAddon);
});

test('does not import intl for Latin or unknown assets', async () => {
  const { loadBundledFontBytes, resolveMetricCompatFace } = await import('./index');
  const face = resolveMetricCompatFace('Arial', false, false)!;
  globalThis.fetch = fontAssetResponse as typeof fetch;
  expect(importIntl).not.toHaveBeenCalled();
  expect((await loadBundledFontBytes(face, { baseUrl: 'https://without-intl.example/' })).byteLength)
    .toBe(face.byteLength);
  await expect(loadBundledFontBytes({ ...face, file: 'Unlisted-Regular.ttf.zst' })).rejects.toThrow(
    /^Unknown bundled font asset: Unlisted-Regular\.ttf\.zst$/,
  );
  expect(importIntl).not.toHaveBeenCalled();
});

test('reports the missing intl add-on for every intl face and retries misses', async () => {
  const { BUNDLED_FONTS, loadBundledFontBytes } = await import('./index');
  let attempts = 0;
  for (const face of BUNDLED_FONTS.filter((candidate) => candidate.file.endsWith('.ttf.zst') && !isLatinFont(candidate.file))) {
    for (const script of [face.script, undefined]) {
      await expect(loadBundledFontBytes({ ...face, script })).rejects.toThrow(
        `Bundled font ${face.file} needs the optional intl add-on - install @betteroffice/fonts-intl`,
      );
      expect(importIntl).toHaveBeenCalledTimes(++attempts);
    }
  }
});

test('recognizes missing-module errors without hiding dependency or evaluation failures', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = resolveScriptFallbackFace('thai', false, false)!;
  for (const message of [
    "Cannot find module '@betteroffice/fonts-intl'",
    'Failed to resolve module specifier "@betteroffice/fonts-intl"',
    'Error resolving module specifier "@betteroffice/fonts-intl"',
    'The specifier "@betteroffice/fonts-intl" could not be resolved',
    "Module name, '@betteroffice/fonts-intl' does not resolve",
  ]) {
    importIntl.mockImplementation(() => { throw new Error(message); });
    await expect(loadBundledFontBytes(face)).rejects.toThrow('install @betteroffice/fonts-intl');
  }
  for (const message of [
    "Cannot find package 'dependency-of-intl'",
    "Cannot find package '@betteroffice/fonts-intl-extra'",
    'Failed to fetch dynamically imported module',
    'intl module evaluation failed',
  ]) {
    const failure = new Error(message);
    importIntl.mockImplementation(() => { throw failure; });
    await expect(loadBundledFontBytes(face)).rejects.toBe(failure);
  }
});

test('rejects invalid intl exports and retries without falling back to Latin', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = resolveScriptFallbackFace('thai', false, false)!;
  let attempts = 0;
  for (const addon of [{}, { INTL_FONT_ASSET_URLS: null }, { INTL_FONT_ASSET_URLS: 'invalid' }]) {
    importIntl.mockImplementation(() => addon);
    for (let retry = 0; retry < 2; retry++) {
      await expect(loadBundledFontBytes(face)).rejects.toThrow(
        /^Invalid INTL_FONT_ASSET_URLS export from @betteroffice\/fonts-intl$/,
      );
      expect(importIntl).toHaveBeenCalledTimes(++attempts);
    }
  }
});

test('custom bases bypass the optional add-on and return owned buffers', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = resolveScriptFallbackFace('mongolian', false, false)!;
  const requested: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requested.push(String(input));
    return fontAssetResponse(input);
  }) as typeof fetch;
  const options = { baseUrl: 'https://without-intl-custom.example/' };
  const first = await loadBundledFontBytes(face, options);
  structuredClone(first, { transfer: [first] });
  expect((await loadBundledFontBytes(face, options)).byteLength).toBe(face.byteLength);
  expect(requested.filter((url) => !url.endsWith('.wasm'))).toEqual([`https://without-intl-custom.example/${face.file}`]);
  expect(importIntl).not.toHaveBeenCalled();
});

test('recovers after installing intl and reports missing asset entries', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = resolveScriptFallbackFace('tifinagh', false, false)!;
  await expect(loadBundledFontBytes(face)).rejects.toThrow('install @betteroffice/fonts-intl');
  importIntl.mockImplementation(() => ({
    INTL_FONT_ASSET_URLS: {
      [face.file]: () => new URL('https://intl-installed.example/NotoSansTifinagh-Regular.ttf.zst'),
    },
  }));
  globalThis.fetch = fontAssetResponse as typeof fetch;
  expect((await loadBundledFontBytes(face)).byteLength).toBe(face.byteLength);
  const other = resolveScriptFallbackFace('thai', false, false)!;
  await expect(loadBundledFontBytes(other)).rejects.toThrow(
    `Unknown bundled font asset: ${other.file}`,
  );
});

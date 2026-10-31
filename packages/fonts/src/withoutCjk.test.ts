import { afterEach, expect, mock, test } from 'bun:test';
import { createHash } from 'node:crypto';
import pack from '../../fonts-cjk/assets/manifest.json';
import { fontAssetResponse } from './test-assets';

type CjkAddon = { CJK_FONT_ASSET_URLS?: unknown };

const missingCjkAddon = (): CjkAddon => {
  throw new Error("Cannot find package '@betteroffice/fonts-cjk'");
};
const importCjk = mock(missingCjkAddon);
mock.module('@betteroffice/fonts-cjk', () => ({
  then(resolve: (addon: CjkAddon) => void, reject: (error: unknown) => void) {
    Promise.resolve().then(importCjk).then(resolve, reject);
  },
}));

afterEach(() => {
  importCjk.mockReset();
  importCjk.mockImplementation(missingCjkAddon);
});

test('missing CJK assets fall back without hiding unknown asset errors', async () => {
  const { BUNDLED_FONTS, loadBundledFontBytes, resolveMetricCompatFace } =
    await import('./index');
  expect(importCjk).not.toHaveBeenCalled();
  const inter = resolveMetricCompatFace('Inter', false, false)!;
  expect((await loadBundledFontBytes(inter)).byteLength).toBe(inter.byteLength);
  await expect(loadBundledFontBytes({ ...inter, file: 'Unlisted-Regular.ttf' })).rejects.toThrow(
    /^Unknown bundled font asset: Unlisted-Regular\.ttf$/
  );
  expect(importCjk).not.toHaveBeenCalled();
  for (const cjk of BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-'))) {
    const fallback = resolveMetricCompatFace(
      cjk.family.includes('Serif') ? 'Times New Roman' : 'Arial', false, false
    )!;
    const bytes = await loadBundledFontBytes(cjk);
    expect(bytes).toEqual(await loadBundledFontBytes(fallback));
    const { script: _script, ...withoutScript } = cjk;
    expect(await loadBundledFontBytes(withoutScript)).toEqual(bytes);
  }
});

test('CJK provider loaders return shipped bytes without changing resolver metadata', async () => {
  const {
    createFontProvider,
    resolveBundledFamilyFace,
    resolveLastResortFace,
    resolveMetricCompatFace,
    resolveMetricCompatFamily,
    resolveScriptFallbackFace,
  } = await import('./index');
  const provider = createFontProvider();

  for (const family of ['Microsoft YaHei', 'Microsoft YaHei Light', 'DengXian', 'Yu Gothic']) {
    for (const [bold, italic] of [[false, false], [true, false], [false, true], [true, true]]) {
      const face =
        resolveMetricCompatFace(family, bold, italic) ??
        resolveLastResortFace(family, bold, italic);
      const load = provider.resolve(family, bold, italic) ?? provider.resolveLastResort(family, bold, italic);
      expect(await load()).toEqual(await provider.resolve('Arial', false, false)!());
      expect(
        resolveMetricCompatFace(family, bold, italic) ??
        resolveLastResortFace(family, bold, italic)
      ).toBe(face);
    }
  }
  expect(resolveMetricCompatFamily('Microsoft YaHei')).toBe('Noto Sans SC');
  expect(resolveBundledFamilyFace('DengXian', true, true)).toBeUndefined();
  expect(resolveScriptFallbackFace('cjk-sc', false, false)?.file).toBe('NotoSansSC-Regular.otf.zst');
  expect(provider.resolveFamily('Noto Sans SC', true, true)).toBeUndefined();
  const sans = await provider.resolve('Arial', false, false)!();
  expect(await provider.resolveFamily('Noto Sans SC', false, false)!()).toEqual(sans);
  expect(await provider.resolveScriptFallback('cjk-sc', true, true)!()).toEqual(sans);
});

test('invalid CJK add-on exports reject without Latin fallback and remain retryable', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = resolveScriptFallbackFace('cjk-sc', false, false)!;
  let attempts = 0;
  for (const addon of [{}, { CJK_FONT_ASSET_URLS: null }, { CJK_FONT_ASSET_URLS: 'invalid' }]) {
    importCjk.mockImplementation(() => addon);
    for (let retry = 0; retry < 2; retry++) {
      await expect(loadBundledFontBytes(face)).rejects.toThrow(
        /^Invalid CJK_FONT_ASSET_URLS export from @betteroffice\/fonts-cjk$/
      );
      expect(importCjk).toHaveBeenCalledTimes(++attempts);
    }
  }
});

test('custom CJK roots bypass the missing add-on and propagate packed asset failures', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = resolveScriptFallbackFace('cjk-sc', false, false)!;
  const realFetch = globalThis.fetch;
  let attempts = 0;
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
    expect(String(input)).toBe(`https://cjk-custom.example/${face.file}`);
    attempts++;
    return attempts === 1 ? new Response(new Uint8Array(4)) : fontAssetResponse(input);
  }) as typeof fetch;
  try {
    const options = { baseUrl: 'https://cjk-custom.example/' };
    await expect(loadBundledFontBytes(face, options)).rejects.toThrow('encoded bytes; expected');
    expect((await loadBundledFontBytes(face, options)).byteLength).toBe(face.byteLength);
    expect(attempts).toBe(2);
    expect(importCjk).not.toHaveBeenCalled();
  } finally {
    globalThis.fetch = realFetch;
  }
});

test('CJK import failures reject without Latin fallback and retry after recovery', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = resolveScriptFallbackFace('cjk-sc', false, false)!;
  const failure = new TypeError('Failed to fetch dynamically imported module');
  importCjk.mockImplementation(() => {
    throw failure;
  });

  await expect(loadBundledFontBytes(face)).rejects.toBe(failure);

  const entry = pack.faces.find((entry) => entry.file === face.file)!;
  const url = new URL('https://cjk-retry.example/NotoSansSC-Regular.otf.zst');
  importCjk.mockImplementation(() => ({
    CJK_FONT_ASSET_URLS: { [face.file]: () => url },
  }));
  const realFetch = globalThis.fetch;
  const fetchCjk = mock(fontAssetResponse);
  globalThis.fetch = fetchCjk as unknown as typeof fetch;
  try {
    const bytes = await loadBundledFontBytes(face);
    expect(bytes.byteLength).toBe(face.byteLength);
    expect(createHash('sha256').update(new Uint8Array(bytes)).digest('hex')).toBe(entry.sha256);
    expect(fetchCjk).toHaveBeenCalledWith(url, undefined);
  } finally {
    globalThis.fetch = realFetch;
  }
});

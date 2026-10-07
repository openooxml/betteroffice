import { expect, mock, test } from 'bun:test';

const importCjk = mock((): { CJK_FONT_ASSET_URLS: Record<string, () => URL> } => {
  throw new Error("Cannot find package '@betteroffice/fonts-cjk'");
});
mock.module('@betteroffice/fonts-cjk', importCjk);

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
    expect(bytes).toBe(await loadBundledFontBytes(fallback));
    const { script: _script, ...withoutScript } = cjk;
    expect(await loadBundledFontBytes(withoutScript)).toBe(bytes);
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
      expect(await load()).toBe(await provider.resolve('Arial', false, false)!());
      expect(
        resolveMetricCompatFace(family, bold, italic) ??
        resolveLastResortFace(family, bold, italic)
      ).toBe(face);
    }
  }
  expect(resolveMetricCompatFamily('Microsoft YaHei')).toBe('Noto Sans SC');
  expect(resolveBundledFamilyFace('DengXian', true, true)).toBeUndefined();
  expect(resolveScriptFallbackFace('cjk-sc', false, false)?.file).toBe('NotoSansSC-Regular.otf');
  expect(provider.resolveFamily('Noto Sans SC', true, true)).toBeUndefined();
  const sans = await provider.resolve('Arial', false, false)!();
  expect(await provider.resolveFamily('Noto Sans SC', false, false)!()).toBe(sans);
  expect(await provider.resolveScriptFallback('cjk-sc', true, true)!()).toBe(sans);
});

test('CJK import failures reject without Latin fallback and retry after recovery', async () => {
  const { loadBundledFontBytes, resolveScriptFallbackFace } = await import('./index');
  const face = { ...resolveScriptFallbackFace('cjk-sc', false, false)!, byteLength: 4 };
  const failure = new TypeError('Failed to fetch dynamically imported module');
  importCjk.mockImplementation(() => {
    throw failure;
  });
  mock.module('@betteroffice/fonts-cjk', importCjk);

  await expect(loadBundledFontBytes(face)).rejects.toBe(failure);

  const expected = new Uint8Array([0x4f, 0x54, 0x54, 0x4f]);
  const url = new URL('https://cjk-retry.example/NotoSansSC-Regular.otf');
  importCjk.mockImplementation(() => ({
    CJK_FONT_ASSET_URLS: { [face.file]: () => url },
  }));
  mock.module('@betteroffice/fonts-cjk', importCjk);
  const realFetch = globalThis.fetch;
  const fetchCjk = mock(async () => new Response(expected));
  globalThis.fetch = fetchCjk as unknown as typeof fetch;
  try {
    expect(await loadBundledFontBytes(face)).toEqual(expected.buffer);
    expect(fetchCjk).toHaveBeenCalledWith(url, undefined);
  } finally {
    globalThis.fetch = realFetch;
  }
});

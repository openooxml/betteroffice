import { expect, mock, test } from 'bun:test';

const importCjk = mock(() => {
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

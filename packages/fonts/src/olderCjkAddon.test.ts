import { expect, mock, test } from 'bun:test';
import pack from '../../fonts-cjk/assets/manifest.json';

const importCjk = mock(() => ({
  CJK_FONT_ASSET_URLS: Object.fromEntries(pack.faces.map((face) => [
    face.source, () => new URL(`https://example.invalid/${face.source}`),
  ])),
}));
mock.module('@betteroffice/fonts-cjk', () => ({
  then(resolve: (addon: unknown) => void, reject: (error: unknown) => void) {
    Promise.resolve().then(importCjk).then(resolve, reject);
  },
}));

test('an older CJK add-on without packed files falls back instead of failing', async () => {
  const { BUNDLED_FONTS, loadBundledFontBytes, resolveMetricCompatFace } = await import('./index');
  for (const cjk of BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-'))) {
    const fallback = resolveMetricCompatFace(
      cjk.family.includes('Serif') ? 'Times New Roman' : 'Arial', false, false
    )!;
    expect(await loadBundledFontBytes(cjk)).toEqual(await loadBundledFontBytes(fallback));
  }
  expect(importCjk).toHaveBeenCalled();
});

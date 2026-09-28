import { expect, mock, test } from 'bun:test';

mock.module('@betteroffice/fonts-cjk', () => {
  throw new Error("Cannot find package '@betteroffice/fonts-cjk'");
});

test('without the CJK add-on only a CJK face asks for it', async () => {
  const { loadBundledFontBytes, resolveMetricCompatFace, resolveScriptFallbackFace } =
    await import('./index');
  const inter = resolveMetricCompatFace('Inter', false, false)!;
  expect((await loadBundledFontBytes(inter)).byteLength).toBe(inter.byteLength);
  await expect(loadBundledFontBytes({ ...inter, file: 'Unlisted-Regular.ttf' })).rejects.toThrow(
    /^Unknown bundled font asset: Unlisted-Regular\.ttf$/
  );
  await expect(
    loadBundledFontBytes(resolveScriptFallbackFace('cjk-sc', false, false)!)
  ).rejects.toThrow('install @betteroffice/fonts-cjk');
  const { script: _script, ...withoutScript } = resolveScriptFallbackFace('cjk-sc', false, false)!;
  await expect(loadBundledFontBytes(withoutScript)).rejects.toThrow('install @betteroffice/fonts-cjk');
});

import { afterEach, expect, mock, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import * as fonts from './index';
import * as worker from './worker';
import { bundledDecoderUrl } from './workerAssets';
import type { FontAssetKind } from '../../../shared/office-session/optionalFontPacks';
import fontsPackage from '../package.json';

afterEach(() => worker.setFontAssetSource(async () => undefined));

test('the worker runtime includes the font and progressive selectors', async () => {
  const progressive = await import('./progressive-entry');
  for (const name of new Set([...Object.keys(fonts), ...Object.keys(progressive)])) {
    if (name === 'fontAssetUrls' || name === 'fontCdnConfiguration') continue;
    expect(Object.prototype.hasOwnProperty.call(worker, name)).toBe(true);
  }
});

test('the worker requests Latin, including base Latin, only when a face is needed', async () => {
  const source = mock(fonts.fontAssetUrls);
  worker.setFontAssetSource(source);
  await worker.checkFontPacks();
  expect(source).not.toHaveBeenCalled();
  const regular = worker.resolveMetricCompatFace('Calibri', false, false)!;
  const [base] = worker.BASE_LATIN_FONTS;
  expect(await Promise.all([worker.bundledFaceInstalled(regular), worker.bundledFaceInstalled(base!)])).toEqual([true, true]);
  expect(source.mock.calls).toEqual([['latin']]);
});

test('intl and cjk faces use the injected URL maps once per pack', async () => {
  const intl = worker.resolveScriptFallbackFace('thai', false, false)!;
  const cjk = worker.resolveScriptFallbackFace('cjk-sc', false, false)!;
  const source = mock(async (pack: FontAssetKind) => {
    const face = pack === 'intl' ? intl : cjk;
    return { [face.file]: `https://fonts.example/${face.file}` };
  });
  worker.setFontAssetSource(source);
  expect(source).not.toHaveBeenCalled();
  for (const face of [intl, intl, cjk, cjk]) expect(await worker.bundledFaceInstalled(face)).toBe(true);
  expect(source.mock.calls.map(([pack]) => pack)).toEqual(['intl', 'cjk']);
});

test('missing packs leave bundled faces unavailable and can be retried', async () => {
  const source = mock(async () => undefined);
  worker.setFontAssetSource(source);
  const latin = worker.resolveMetricCompatFace('Calibri', false, false)!;
  expect(await worker.bundledFaceInstalled(latin)).toBe(false);
  await expect(worker.loadBundledFontBytes(latin)).rejects.toThrow(latin.file);
  expect(source.mock.calls).toEqual([['latin'], ['latin']]);
});

test('worker pack lookups coalesce failures and retry without resolving URLs against the worker', async () => {
  const face = worker.resolveScriptFallbackFace('thai', false, false)!;
  const href = `https://cdn.example/${face.file}`;
  const hrefs = { [face.file]: href };
  const failure = new Error('Pack source unavailable');
  const source = mock(async (): Promise<Record<string, string> | undefined> => hrefs)
    .mockRejectedValueOnce(failure).mockResolvedValueOnce(undefined);
  worker.setFontAssetSource(source);
  const replies = await Promise.allSettled([worker.bundledFaceInstalled(face), worker.bundledFaceInstalled(face)]);
  expect(replies).toEqual([{ status: 'rejected', reason: failure }, { status: 'rejected', reason: failure }]);
  expect(source).toHaveBeenCalledTimes(1);
  expect(await worker.bundledFaceInstalled(face)).toBe(false);
  expect(await worker.bundledFaceInstalled(face)).toBe(true);
  expect(await worker.bundledFaceInstalled(face)).toBe(true);
  expect(source).toHaveBeenCalledTimes(3);
});

test('the decoder URL comes from its own request and is absent when that request returns none', async () => {
  const href = 'https://page.example/assets/published-decoder.wasm';
  const source = mock(fonts.fontAssetUrls).mockResolvedValueOnce(undefined).mockResolvedValueOnce({ 'fontpack.wasm': href });
  worker.setFontAssetSource(source);
  expect(await bundledDecoderUrl()).toBeUndefined();
  expect((await bundledDecoderUrl())?.href).toBe(href);
  expect(source.mock.calls).toEqual([['decoder'], ['decoder']]);
});

test('the published main entry exposes the page asset lookup and recognizes a worker missing-file error', async () => {
  expect(Object.keys(fontsPackage.exports)).not.toContain('./worker');
  const types = await readFile(new URL('../dist/index.d.ts', import.meta.url), 'utf8');
  expect(types).toContain('fontAssetUrls');
  const main: typeof fonts = await import(new URL('../dist/index.js', import.meta.url).href);
  expect(await main.fontAssetUrls('decoder')).toEqual({ 'fontpack.wasm': new URL('../dist/decoder/fontpack.wasm', import.meta.url).href });
  const face = worker.resolveScriptFallbackFace('cjk-sc', false, false)!;
  worker.setFontAssetSource(async (kind) => kind === 'cjk' ? { 'other-face.zst': 'https://fonts.example/other-face.zst' } : main.fontAssetUrls(kind));
  expect(await worker.bundledFaceInstalled(face)).toBe(false);
  const error: unknown = await worker.loadBundledFontBytes(face).catch((cause: unknown) => cause);
  expect(main.isMissingFontAddon(error)).toBe(true);
  expect(error).toMatchObject({ message: expect.stringContaining(face.file) });
  const location = worker.createFontProvider().resolveScriptFallback('cjk-sc', false, false)!.slices!;
  await expect(location.url(location.ref.index)).rejects.toThrow(location.ref.index);
});

import { afterEach, expect, mock, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import * as fonts from './index';
import * as worker from './worker';
import { isLatinFont } from './packed';
import fontsPackage from '../package.json';

afterEach(() => worker.setOptionalFontPacks(async () => undefined));

test('the worker entry includes the runtime exports of both font entries', async () => {
  const progressive = await import('./progressive-entry');
  for (const name of new Set([...Object.keys(fonts), ...Object.keys(progressive)])) {
    expect(Object.prototype.hasOwnProperty.call(worker, name)).toBe(true);
  }
});

test('the worker entry resolves installed Latin faces like the default entry without requesting optional packs', async () => {
  const source = mock(async () => undefined);
  worker.setOptionalFontPacks(source);
  for (const face of fonts.BUNDLED_FONTS.filter((face) => isLatinFont(face.file))) {
    expect(worker.resolveBundledFamilyFace(face.family, face.weight === 700, face.style === 'italic'))
      .toEqual(fonts.resolveBundledFamilyFace(face.family, face.weight === 700, face.style === 'italic'));
    expect(await worker.bundledFaceInstalled(face)).toBe(await fonts.bundledFaceInstalled(face));
  }
  await worker.checkFontPacks();
  expect(source).not.toHaveBeenCalled();
});

test('intl and cjk faces use the injected URL maps once per pack', async () => {
  const intl = worker.resolveScriptFallbackFace('thai', false, false)!;
  const cjk = worker.resolveScriptFallbackFace('cjk-sc', false, false)!;
  const source = mock(async (pack: worker.OptionalFontPack) => {
    const face = pack === 'intl' ? intl : cjk;
    return { [face.file]: `https://fonts.example/${face.file}` };
  });
  worker.setOptionalFontPacks(source);
  expect(source).not.toHaveBeenCalled();
  for (const face of [intl, intl, cjk, cjk]) expect(await worker.bundledFaceInstalled(face)).toBe(true);
  expect(source.mock.calls.map(([pack]) => pack)).toEqual(['intl', 'cjk']);
});

test('missing injected packs keep optional faces unavailable', async () => {
  const source = mock(async () => undefined);
  worker.setOptionalFontPacks(source);
  const face = worker.resolveScriptFallbackFace('cjk-sc', false, false)!;
  expect(await worker.bundledFaceInstalled(face)).toBe(false);
  expect(await worker.bundledFaceInstalled(face)).toBe(false);
  expect(source).toHaveBeenCalledTimes(2);
});

test('worker pack URLs retry failures and misses and resolve hrefs from the injected source', async () => {
  const hrefs = { 'path.zst': '/_next/static/media/path.zst', 'cdn.zst': 'https://cdn.example/cdn.zst' };
  const source = mock(async (): Promise<Record<string, string> | undefined> => hrefs)
    .mockRejectedValueOnce(new Error('Pack source unavailable')).mockResolvedValueOnce(undefined);
  worker.setOptionalFontPacks(source);
  const location = Object.getOwnPropertyDescriptor(globalThis, 'location');
  Object.defineProperty(globalThis, 'location', { value: { href: 'https://worker.example/room/1' }, configurable: true });
  try {
    const first = worker.optionalFontPackUrls('intl');
    const concurrent = worker.optionalFontPackUrls('intl');
    await Promise.all([first, concurrent].map((request) => expect(request).rejects.toThrow('Pack source unavailable')));
    expect(source).toHaveBeenCalledTimes(1);
    expect(await worker.optionalFontPackUrls('intl')).toBeUndefined();
    const absolute = { ...hrefs, 'path.zst': 'https://worker.example/_next/static/media/path.zst' };
    expect(await worker.optionalFontPackUrls('intl')).toEqual(absolute);
    expect(await worker.optionalFontPackUrls('intl')).toEqual(absolute);
    expect(source).toHaveBeenCalledTimes(3);
  } finally {
    if (location) Object.defineProperty(globalThis, 'location', location);
    else Reflect.deleteProperty(globalThis, 'location');
  }
});

test('the published worker entry has types and no literal dynamic imports', async () => {
  expect(fontsPackage.exports['./worker']).toEqual({ types: './dist/worker.d.ts', import: './dist/worker.js', default: './dist/worker.js' });
  const source = await readFile(new URL('../dist/worker.js', import.meta.url), 'utf8');
  expect(source).not.toMatch(/\bimport\s*\(\s*['"`]/);
  expect(source).toMatch(/from\s+['"]@betteroffice\/fonts-latin['"]/);
  expect(source).not.toMatch(/(?:from\s*|import\s*\()\s*['"]@betteroffice\/fonts-(?:intl|cjk)['"]/);
  const types = await readFile(new URL('../dist/worker.d.ts', import.meta.url), 'utf8');
  expect(types).toContain('setOptionalFontPacks');
  expect(types).toContain('optionalFontPackUrls');
});

test('the published entries recognize each other\'s missing-pack errors', async () => {
  const main: typeof fonts = await import(new URL('../dist/index.js', import.meta.url).href);
  const entry: typeof worker = await import(new URL('../dist/worker.js', import.meta.url).href);
  entry.setOptionalFontPacks(async () => undefined);
  const face = entry.resolveScriptFallbackFace('cjk-sc', false, false)!;
  const error: unknown = await entry.loadBundledFontBytes(face, undefined, { strict: true }).catch((cause: unknown) => cause);
  expect(main.isMissingFontAddon(error)).toBe(true);
});

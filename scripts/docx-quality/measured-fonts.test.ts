import { expect, test } from 'bun:test';
import type { SliceLocation } from '../../packages/fonts/src/progressive';
import { fontProvider, type BundledFontLoader, type BundledFontSource } from '../../packages/fonts/src/provider';
import { measuredFontProvider, type FontLoad } from './measured-fonts';

test('the measured provider forwards and records glyph coverage loads', async () => {
  const provider = fontProvider(async (face) => {
    if (face.file === 'Carlito-Bold.ttf.zst') throw new Error('offline');
    return new ArrayBuffer(face.byteLength);
  });
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(provider, fontLoads);
  expect(measured.resolve('Segoe UI', false, false)!.measureAs).toBe('Segoe UI');
  expect(measured.resolveGlyphFallback).toBeDefined();
  expect(measured.resolveGlyphFallback!('Google Sans', false, false)).toBeUndefined();
  const coverage = measured.resolveGlyphFallback!('Segoe UI Black', true, false, 'word')!;
  expect((await coverage()).byteLength).toBe(
    (await provider.resolveGlyphFallback!('Segoe UI Black', true, false, 'word')!()).byteLength,
  );
  await expect(measured.resolveGlyphFallback!('Segoe UI Black', true, false, 'powerpoint')!()).rejects.toThrow('offline');
  expect(fontLoads).toEqual([
    { kind: 'glyph', family: 'Segoe UI Black', ok: true },
    { kind: 'glyph', family: 'Segoe UI Black', ok: false, error: 'Error: offline' },
  ]);
});

const fakeLoader = (file: string, bytes: ArrayBuffer, slices?: SliceLocation): BundledFontLoader => {
  const calls = { load: 0, readOnly: 0 };
  return Object.assign(
    async () => {
      calls.load++;
      return bytes.slice(0);
    },
    {
      calls,
      readOnlyBytes: async () => {
        calls.readOnly++;
        if (file === 'Missing.ttf.zst') throw new Error('offline');
        return bytes;
      },
    },
    slices ? { slices } : {},
  );
};

const fakeSlices = (base = 'https://fonts.example/', url?: SliceLocation['url']): SliceLocation => ({
  ref: { format: 1, index: 'cjk/NotoSansJP/slices.json', indexSha256: '00' } as SliceLocation['ref'],
  base,
  url: url ?? (async (file) => new URL(file, base)),
  decode: async (encoded) => encoded,
});

const fakeSource = (loaders: Record<string, BundledFontLoader>): BundledFontSource => ({
  resolve: (family) => loaders[family],
  resolveFamily: (family) => loaders[`bundled ${family}`],
  resolveScriptFallback: (script) => loaders[script],
  resolveLastResort: () => loaders.fallback!,
  resolveLightVariant: (family) => loaders[`light ${family}`],
});

test('the measured provider records read-only byte loads', async () => {
  const latin = new ArrayBuffer(8);
  const loaders = {
    Calibri: fakeLoader('Carlito-Regular.ttf.zst', latin),
    fallback: fakeLoader('Missing.ttf.zst', new ArrayBuffer(4)),
  };
  const files: Record<string, string> = { family: 'Carlito-Regular.ttf.zst', fallback: 'Missing.ttf.zst' };
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(fakeSource(loaders), fontLoads, (kind) => files[kind]);
  const fetch = (loader: BundledFontLoader) => (loader.readOnlyBytes ?? loader)();

  expect(await fetch(measured.resolve('Calibri', false, false)!)).toBe(latin);
  await expect(fetch(measured.resolveLastResort('Wingdings', false, false))).rejects.toThrow('offline');

  expect(fontLoads).toEqual([
    { kind: 'family', family: 'Calibri', ok: true, face: 'Carlito-Regular.ttf.zst' },
    { kind: 'fallback', family: 'Wingdings', ok: false, error: 'Error: offline', face: 'Missing.ttf.zst' },
  ]);
  for (const loader of Object.values(loaders))
    expect((loader as BundledFontLoader & { calls: { load: number } }).calls.load).toBe(0);
});

test('a sliced loader passes its slices through untouched and slice urls record nothing', async () => {
  const slices = fakeSlices();
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(
    fakeSource({ 'cjk-jp': fakeLoader('NotoSansJP-Regular.otf.zst', new ArrayBuffer(16), slices) }),
    fontLoads,
    () => 'NotoSansJP-Regular.otf.zst',
  );
  const loader = measured.resolveScriptFallback('cjk-jp', false, false)!;
  expect(loader.slices).toBe(slices);
  for (const file of ['cjk/NotoSansJP/slices.json', 'cjk/NotoSansJP/skeleton.zst', 'cjk/NotoSansJP/chunk-1.zst'])
    expect((await loader.slices!.url(file)).href).toBe(`https://fonts.example/${file}`);
  expect(fontLoads).toEqual([]);
});

test('a synchronous throw from a loader or its read-only bytes is recorded as a failure and rethrown unchanged', async () => {
  const boom = new Error('no base');
  const throwing = Object.assign(
    () => {
      throw boom;
    },
    {
      readOnlyBytes: () => {
        throw boom;
      },
    },
  ) as unknown as BundledFontLoader;
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(fakeSource({ Calibri: throwing }), fontLoads);
  const caught = async (run: () => Promise<ArrayBuffer>) => {
    try {
      await run();
    } catch (error) {
      return error;
    }
  };
  const loader = measured.resolve('Calibri', false, false)!;
  expect(await caught(() => loader.readOnlyBytes!())).toBe(boom);
  expect(await caught(() => loader())).toBe(boom);
  expect(fontLoads).toEqual([
    { kind: 'family', family: 'Calibri', ok: false, error: 'Error: no base' },
    { kind: 'family', family: 'Calibri', ok: false, error: 'Error: no base' },
  ]);
});

test('the measured provider forwards and records bundled family and light variant loads', async () => {
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(
    fakeSource({
      'bundled Gelasio': fakeLoader('Gelasio-Regular.ttf.zst', new ArrayBuffer(2)),
      'light Calibri Light': fakeLoader('Carlito-Regular.ttf.zst', new ArrayBuffer(2)),
    }),
    fontLoads,
  );
  expect(measured.resolveFamily!('Missing', false, false)).toBeUndefined();
  await measured.resolveFamily!('Gelasio', false, false)!.readOnlyBytes!();
  await measured.resolveLightVariant!('Calibri Light', false, false)!();
  expect(fontLoads).toEqual([
    { kind: 'bundled-family', family: 'Gelasio', ok: true },
    { kind: 'light-variant', family: 'Calibri Light', ok: true },
  ]);
});

test('the measured provider leaves glyph coverage out when the source has none', () => {
  const { resolveGlyphFallback: _omitted, ...source } = fontProvider(async () => new ArrayBuffer(0));
  expect('resolveGlyphFallback' in measuredFontProvider(source, [])).toBe(false);
});

import { expect, test } from 'bun:test';
import { createFontProvider as createCdnFontProvider } from './cdn';
import { createFontProvider } from './index';
import { fontProvider } from './provider';

for (const [source, create] of [
  ['bundled', createFontProvider],
  ['CDN', createCdnFontProvider],
] as const) {
  for (const family of ['Segoe UI Light', 'Segoe UI Semilight', 'Segoe UI', 'Segoe UI Semibold']) {
    for (const bold of [false, true]) {
      test(`${source} ${family} ${bold} shares its Word coverage and last-resort loader`, () => {
        const provider = create();
        const lastResort = provider.resolveLastResort(family, bold, false, 'word');
        const coverage = provider.resolveGlyphFallback!(family, bold, false, 'word');
        const coverageBold = / (Light|Semilight)$/.test(family) ? false : bold;
        expect(coverage).toBeDefined();
        expect(lastResort).toBe(coverage!);
        expect(provider.resolve('Arial', coverageBold, false)).toBe(coverage);
        expect(provider.resolveFamily('Liberation Sans', coverageBold, false)).toBe(coverage);
        expect(lastResort.onlyIfPrimaryFails).toBeUndefined();
        expect(provider.resolve(family, bold, false)!.measureAs)
          .toBe(family === 'Segoe UI Semibold' && bold ? 'Segoe UI' : family);
        expect(coverage).not.toBe(provider.resolveGlyphFallback!(family, bold, false, 'powerpoint'));
        expect(coverage).not.toBe(create().resolveGlyphFallback!(family, bold, false, 'word'));
      });
    }
  }
}

test('shared loaders stay lazy and return independently owned buffers after transfer or mutation', async () => {
  const loaded: string[] = [];
  const bytes = new Uint8Array([0, 1, 0, 0, 42]);
  const provider = fontProvider(async (face) => {
    loaded.push(face.file);
    return bytes.slice().buffer;
  });
  const lastResort = provider.resolveLastResort('Segoe UI', false, false, 'word');
  const coverage = provider.resolveGlyphFallback!('Segoe UI', false, false, 'word')!;
  expect(lastResort).toBe(coverage);
  expect(loaded).toEqual([]);
  const [first, second] = await Promise.all([lastResort(), coverage()]);
  expect(first).not.toBe(second);
  expect(new Uint8Array(first)).toEqual(bytes);
  expect(new Uint8Array(second)).toEqual(bytes);
  structuredClone(first, { transfer: [first] });
  new Uint8Array(second).fill(0);
  expect(new Uint8Array(await coverage())).toEqual(bytes);
  expect(loaded).toEqual(Array(3).fill('LiberationSans-Regular.ttf.zst'));
});

test('shared face loaders keep conditional fallback and measurement metadata separate', () => {
  const provider = fontProvider(async () => new ArrayBuffer(16));
  const primary = provider.resolve('Calibri', false, false)!;
  const conditional = provider.resolveLastResort('Arial', false, false, 'word');
  expect(conditional.onlyIfPrimaryFails).toBe(true);
  expect(primary.onlyIfPrimaryFails).toBeUndefined();
  expect(conditional).not.toBe(primary);
  expect(provider.resolveLastResort('Arial', false, false, 'word')).toBe(conditional);
  expect(provider.resolve('Calibri', false, false)).toBe(primary);
  const measured = provider.resolve('Segoe UI', false, false)!;
  const raw = provider.resolveFamily('Selawik', false, false)!;
  expect(measured.measureAs).toBe('Segoe UI');
  expect(raw).not.toBe(measured);
  expect(provider.resolveFamily('Selawik', false, false)).toBe(raw);
  expect(provider.resolve('Segoe UI', false, false)).toBe(measured);
});

import { afterEach, expect, spyOn, test } from 'bun:test';
import { DEFAULT_COMPATIBILITY_FLAGS } from '../../docx/settingsParser';
import { configureDefaultFonts, resolveDefaultFontProvider } from './defaultFontProvider';
import {
  createRustMeasureSource,
  type ResidentFontRequirement,
} from './rustMeasureSource';

afterEach(() => configureDefaultFonts({}));

test('default providers retain the decoder of each editor runtime', async () => {
  const first = { decoder: async () => ({ font: new ArrayBuffer(4) }) };
  const second = { decoder: async () => ({ font: new ArrayBuffer(8) }) };
  const seen: unknown[] = [];
  configureDefaultFonts({ fonts: { createFontProvider(_options, runtime) {
    seen.push(runtime);
    return { resolve: () => async () => (await runtime!.decoder!(new ArrayBuffer(0), { codec: 'zstd' })).font };
  } } });
  const a = await resolveDefaultFontProvider(first);
  const b = await resolveDefaultFontProvider(second);
  expect(await resolveDefaultFontProvider(first)).toBe(a);
  expect(a).not.toBe(b);
  expect((await a!.resolve('Arial', false, false)!()).byteLength).toBe(4);
  expect((await b!.resolve('Arial', false, false)!()).byteLength).toBe(8);
  expect(seen).toEqual([first, second]);
});

test('a Japanese preflight settles atomically when font sources are unavailable', async () => {
  configureDefaultFonts({
    load: async () => {
      throw new Error('offline');
    },
  });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const source = createRustMeasureSource({
      engine: { registerFont: () => 1, clearFonts() {} },
    });
    const requirements: ResidentFontRequirement[] = [
      'Century',
      'MS Mincho',
      'Calibri',
    ].map((family) => ({
      key: `${family}|0|0`,
      family,
      bold: false,
      italic: false,
      scripts: ['cjk-jp'],
    }));
    expect(
      source.measurementConfigForRequirements(requirements),
    ).toBeUndefined();
    await source.prepareFontRequirements(requirements);
    const settled = source.measurementConfigForRequirements(requirements);
    expect(settled).toBeDefined();
    expect(settled?.fontChains).toEqual({});
    await Promise.resolve();
    await source.prepareFontRequirements(requirements);
    expect(source.measurementConfigForRequirements(requirements)).toBeDefined();
  } finally {
    warn.mockRestore();
  }
});

test('partial cache reads do not consume a failed family before other fonts settle', async () => {
  let release: ((bytes: ArrayBuffer) => void) | undefined;
  const pending = new Promise<ArrayBuffer>((resolve) => {
    release = resolve;
  });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const source = createRustMeasureSource({
      engine: { registerFont: () => 7, clearFonts() {} },
      bundled: {
        resolve: (family) =>
          family === 'Slow'
            ? () => pending
            : async () => {
                throw new Error('missing');
              },
      },
    });
    const fast = { key: 'fast', family: 'Fast', bold: false, italic: false };
    const slow = { key: 'slow', family: 'Slow', bold: false, italic: false };
    await source.prepareFontRequirements([fast]);
    expect(
      source.measurementConfigForRequirements([fast, slow]),
    ).toBeUndefined();
    await Promise.resolve();
    const preparing = source.prepareFontRequirements([slow]);
    release!(new ArrayBuffer(4));
    await preparing;
    expect(
      source.measurementConfigForRequirements([fast, slow])?.fontChains,
    ).toEqual({ slow: [7] });
  } finally {
    warn.mockRestore();
  }
});

test('rejecting font loads settle to fallback metrics and warn once across passes', async () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const reject = async (): Promise<ArrayBuffer> => {
    throw new Error('font CDN blocked');
  };
  const source = createRustMeasureSource({
    engine: { registerFont: () => { throw new Error('unexpected font bytes'); }, clearFonts() {} },
    bundled: { resolve: () => reject, resolveLastResort: () => reject },
  });
  const requirements = ['Arial', 'Calibri'].map(family => ({
    key: `${family.toLowerCase()}|0|0`, family, bold: false, italic: false,
  }));
  try {
    for (let pass = 0; pass < 2; pass++) {
      await source.prepareFontRequirements(requirements);
      expect(source.measurementConfigForRequirements(requirements)?.fontChains).toEqual({});
      await Promise.resolve();
    }
    expect(warn).toHaveBeenCalledTimes(1);
    expect(String(warn.mock.calls[0]![0])).toContain('falling back');
  } finally {
    source.clear();
    warn.mockRestore();
  }
});

test('measurement forwards the document compatibility mode and defaults to 12', () => {
  const source = createRustMeasureSource({
    engine: { registerFont: () => 1, clearFonts() {} },
  });
  expect(source.measurementConfigForRequirements([])?.compat.compatibilityMode).toBe(12);
  for (const compatibilityMode of [11, 12, 14, 15, 16]) {
    source.setCompat({ ...DEFAULT_COMPATIBILITY_FLAGS, compatibilityMode });
    expect(source.measurementConfigForRequirements([])?.compat.compatibilityMode).toBe(
      compatibilityMode,
    );
  }
});

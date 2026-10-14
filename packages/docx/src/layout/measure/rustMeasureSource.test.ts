import { afterEach, expect, spyOn, test } from 'bun:test';
import { configureDefaultFonts } from './defaultFontProvider';
import {
  createRustMeasureSource,
  type ResidentFontRequirement,
} from './rustMeasureSource';

afterEach(() => configureDefaultFonts({}));

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

import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import * as fontManager from '@betteroffice/fonts';
import { DEFAULT_COMPATIBILITY_FLAGS } from '../../docx/settingsParser';
import { configureDefaultFonts, IMPLICIT_FONT_COVERAGE } from './defaultFontProvider';
import {
  createRustMeasureSource,
  type ResidentFontRequirement,
} from './rustMeasureSource';

afterEach(() => configureDefaultFonts({}));

function providerFor(data: ArrayBuffer) {
  const load = async () => data;
  return {
    resolve: mock(() => load),
    resolveFamily: () => load,
    resolveScriptFallback: () => undefined,
    resolveLastResort: () => load,
  };
}

const regular: ResidentFontRequirement = {
  key: 'calibri|0|0', family: 'Calibri', bold: false, italic: false,
};

test('implicit measurement leaves unknown families absent even with script fallbacks', async () => {
  const registerFont = mock((_bytes: Uint8Array) => 7);
  const source = createRustMeasureSource({ engine: { registerFont, clearFonts() {} } });
  const unknown: ResidentFontRequirement = {
    key: 'brand sans|0|0', family: 'Brand Sans', bold: false, italic: false,
  };
  await source.prepareFontRequirements([unknown]);
  expect(source.measurementConfigForRequirements([unknown])?.fontChains).toEqual({});
  expect(registerFont).not.toHaveBeenCalled();
  unknown.scripts = ['arabic'];
  await source.prepareFontRequirements([unknown]);
  expect(source.measurementConfigForRequirements([unknown])?.fontChains).toEqual({});
  const liberation = await fontManager.createFontProvider().resolveLastResort('Brand Sans', false, false, 'word')();
  expect(registerFont.mock.calls.map(([bytes]) => bytes)).not.toContainEqual(new Uint8Array(liberation));
  await source.prepareFontRequirements([regular]);
  const carlito = await fontManager.createFontProvider().resolve('Calibri', false, false)!();
  expect(registerFont).toHaveBeenCalledWith(new Uint8Array(carlito));
});

test('unconfigured measurement lazily resolves faces through the fonts manager', async () => {
  const data = new Uint8Array([1, 2, 3]).buffer;
  const provider = providerFor(data);
  const create = spyOn(fontManager, 'createFontProvider').mockReturnValue(provider);
  const registerFont = mock(() => 7);
  try {
    const source = createRustMeasureSource({ engine: { registerFont, clearFonts() {} } });
    expect(create).not.toHaveBeenCalled();
    expect(source.measurementConfigForRequirements([regular])).toBeUndefined();
    await source.prepareFontRequirements([regular]);
    expect(create).toHaveBeenCalledTimes(1);
    expect(provider.resolve).toHaveBeenCalledWith('Calibri', false, false);
    expect(registerFont).toHaveBeenCalledWith(new Uint8Array(data));
    expect(source.measurementConfigForRequirements([regular])?.fontChains).toEqual({
      [regular.key]: [7],
      [IMPLICIT_FONT_COVERAGE]: [7],
    });
  } finally { create.mockRestore(); }
});

for (const configured of ['module', 'loader', 'explicit'] as const) {
  test(`${configured} font configuration wins over the fonts manager`, async () => {
    const data = new Uint8Array([4, 5, 6]).buffer;
    const provider = providerFor(data);
    const module = { createFontProvider: mock(() => provider) };
    const loader = mock(async () => module);
    const create = spyOn(fontManager, 'createFontProvider').mockImplementation(() => {
      throw new Error('unexpected manager default');
    });
    const registerFont = mock(() => 9);
    try {
      configureDefaultFonts(configured === 'module' ? { fonts: module, load: loader } : { load: loader });
      const source = createRustMeasureSource({
        engine: { registerFont, clearFonts() {} },
        ...(configured === 'explicit' && { bundled: provider }),
      });
      await source.prepareFontRequirements([regular]);
      expect(source.measurementConfigForRequirements([regular])?.fontChains).toEqual({
        [regular.key]: [9],
      });
      expect(registerFont).toHaveBeenCalledWith(new Uint8Array(data));
      expect(create).not.toHaveBeenCalled();
      expect(loader).toHaveBeenCalledTimes(configured === 'loader' ? 1 : 0);
      expect(module.createFontProvider).toHaveBeenCalledTimes(configured === 'explicit' ? 0 : 1);
    } finally { create.mockRestore(); }
  });
}

for (const configured of ['default', 'configured', 'explicit'] as const) {
  test(`embedded faces precede the ${configured} font provider`, async () => {
    const embedded = new Uint8Array([7, 8]).buffer;
    const bundled = new Uint8Array([9, 10]).buffer;
    const provider = providerFor(bundled);
    const create = spyOn(fontManager, 'createFontProvider').mockReturnValue(provider);
    const registered: Uint8Array[] = [];
    try {
      if (configured === 'configured') configureDefaultFonts({ fonts: { createFontProvider: () => provider } });
      const source = createRustMeasureSource({
        engine: {
          registerFont(bytes) { registered.push(bytes); return registered.length; },
          clearFonts() {},
        },
        ...(configured === 'explicit' && { bundled: provider }),
      });
      source.setEmbeddedFaces([{
        family: 'Calibri', weight: 'normal', style: 'normal', data: embedded, subsetted: false,
      }]);
      await source.prepareFontRequirements([regular]);
      expect(registered).toEqual([new Uint8Array(embedded), new Uint8Array(bundled)]);
      expect(source.measurementConfigForRequirements([regular])?.fontChains).toEqual({
        [regular.key]: [1, 2],
        ...(configured === 'default' && { [IMPLICIT_FONT_COVERAGE]: [1, 2] }),
      });
      expect(create).toHaveBeenCalledTimes(configured === 'default' ? 1 : 0);
    } finally { create.mockRestore(); }
  });
}

test('a failed default manager load settles to approximate metrics and retries', async () => {
  const provider = providerFor(new ArrayBuffer(4));
  const failure = new Error('manager unavailable');
  const create = spyOn(fontManager, 'createFontProvider')
    .mockImplementationOnce(() => { throw failure; })
    .mockReturnValue(provider);
  const registerFont = mock(() => 3);
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const source = createRustMeasureSource({ engine: { registerFont, clearFonts() {} } });
    await source.prepareFontRequirements([regular]);
    expect(source.measurementConfigForRequirements([regular])?.fontChains).toEqual({});
    expect(registerFont).not.toHaveBeenCalled();
    expect(warn).toHaveBeenCalledWith(
      '[configureDefaultFonts] the configured font source failed to load', failure,
    );
    expect(warn.mock.calls.some(([message]) => String(message).includes('approximate fallback metrics'))).toBe(true);
    await Promise.resolve();
    await source.prepareFontRequirements([regular]);
    expect(source.measurementConfigForRequirements([regular])?.fontChains).toEqual({
      [regular.key]: [3],
      [IMPLICIT_FONT_COVERAGE]: [3],
    });
    expect(create).toHaveBeenCalledTimes(2);
    expect(warn).toHaveBeenCalledTimes(2);
  } finally { create.mockRestore(); warn.mockRestore(); }
});

test('a failed default manager face remains retryable', async () => {
  let unavailable = true;
  const data = new ArrayBuffer(4);
  const load = mock(async () => {
    if (unavailable) throw new Error('face unavailable');
    return data;
  });
  const create = spyOn(fontManager, 'createFontProvider').mockReturnValue({
    resolve: () => load, resolveFamily: () => load,
    resolveScriptFallback: () => undefined, resolveLastResort: () => load,
  });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const source = createRustMeasureSource({ engine: { registerFont: () => 5, clearFonts() {} } });
    await source.prepareFontRequirements([regular]);
    expect(source.measurementConfigForRequirements([regular])?.fontChains).toEqual({});
    expect(warn).toHaveBeenCalledTimes(1);
    expect(String(warn.mock.calls[0]![0])).toContain('falling back');
    unavailable = false;
    await Promise.resolve();
    await source.prepareFontRequirements([regular]);
    expect(source.measurementConfigForRequirements([regular])?.fontChains).toEqual({
      [regular.key]: [5],
      [IMPLICIT_FONT_COVERAGE]: [5],
    });
    expect(create).toHaveBeenCalledTimes(1);
    expect(warn).toHaveBeenCalledTimes(1);
  } finally { create.mockRestore(); warn.mockRestore(); }
});

test('a base URL configures the default manager without a source warning', async () => {
  const create = spyOn(fontManager, 'createFontProvider').mockReturnValue(providerFor(new ArrayBuffer(4)));
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    configureDefaultFonts({ baseUrl: 'https://example.test/fonts/' });
    expect(create).not.toHaveBeenCalled();
    const source = createRustMeasureSource({ engine: { registerFont: () => 1, clearFonts() {} } });
    await source.prepareFontRequirements([regular]);
    expect(create).toHaveBeenCalledWith({ baseUrl: new URL('https://example.test/fonts/') });
    expect(warn).not.toHaveBeenCalled();
  } finally { create.mockRestore(); warn.mockRestore(); }
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

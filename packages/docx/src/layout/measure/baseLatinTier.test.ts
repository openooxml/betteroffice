import { expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as fontManager from '@betteroffice/fonts';
import { TextMeasureFontRegistry, type BundledFontProvider } from './fontRegistry';
import { createRustMeasureSource } from './rustMeasureSource';
import { configureDefaultFonts } from './defaultFontProvider';
import { charactersFromBits, coverageBits, fontCoverage } from './fontCoverage';

function face(path: string): ArrayBuffer {
  return Uint8Array.from(readFileSync(new URL(`../../../../../${path}`, import.meta.url))).buffer;
}

const regular = face('crates/base-latin/fonts/BetterOfficeBaseLatin-Regular.ttf');
const primary = face('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const hebrew = face('packages/fonts/assets/NotoSansHebrew-Regular.ttf');

test('hosted fonts without registered app bytes keep the implicit chain absent', async () => {
  configureDefaultFonts({});
  const load = mock(async () => regular);
  const create = spyOn(fontManager, 'createFontProvider').mockReturnValue({
    resolve: () => async () => primary, resolveFamily: () => async () => primary,
    resolveScriptFallback: () => undefined, resolveLastResort: () => async () => primary,
    resolveBaseLatin: () => load,
  });
  const bundled = Object.assign({ resolve: () => undefined }, {
    [Symbol.for('@betteroffice/docx/hostedFontFamilies')]: new Set(['requested']),
  });
  try {
    const registerFont = mock(() => 0);
    const source = createRustMeasureSource({ engine: { registerFont, clearFonts() {} }, bundled });
    source.setEmbeddedFaces([]);
    const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['A'] };
    await source.prepareFontRequirements([requirement]);
    expect(source.measurementConfigForRequirements([requirement])?.fontChains).toEqual({});
    expect(source.measurementConfigForRequirements([{ ...requirement, characters: ['A', 'B'] }])?.fontChains).toEqual({});
    expect(registerFont).not.toHaveBeenCalled();
    expect(load).not.toHaveBeenCalled();
  } finally { create.mockRestore(); configureDefaultFonts({}); }
});

test('newly seen covered text settles synchronously with the existing chain', async () => {
  const load = mock(async () => regular);
  const source = createRustMeasureSource({ engine: { registerFont: () => 0, clearFonts() {} },
    bundled: { resolve: () => async () => primary, resolveBaseLatin: () => load } });
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['A'] };
  await source.prepareFontRequirements([requirement]);
  const config = source.measurementConfigForRequirements([requirement]);
  expect(config).toBeDefined();
  expect(source.measurementConfigForRequirements([{ ...requirement, characters: ['A', 'B'] }])).toEqual(config);
  expect(source.measurementConfigForRequirements([{ ...requirement, characters: ['A', 'B', '日'] }])).toEqual(config);
  expect(load).not.toHaveBeenCalled();
});

test('a prepared backstop covers new candidates synchronously', async () => {
  const load = mock(async () => regular);
  let next = 0;
  const source = createRustMeasureSource({ engine: { registerFont: () => next++, clearFonts() {} },
    bundled: { resolve: () => async () => hebrew, resolveBaseLatin: () => load } });
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['Ω'] };
  await source.prepareFontRequirements([requirement]);
  const config = source.measurementConfigForRequirements([requirement]);
  expect(source.measurementConfigForRequirements([{ ...requirement, characters: ['Ω', 'Ж'] }])).toEqual(config);
  expect(config?.fontChains[requirement.key]).toEqual([0, 1]);
  expect(load).toHaveBeenCalledTimes(1);
});

test('aliased requirement keys settle with independent character sets', async () => {
  let next = 0;
  const source = createRustMeasureSource({
    engine: { registerFont: () => next++, clearFonts() {} },
    bundled: { resolve: () => async () => primary, resolveBaseLatin: () => async () => regular },
  });
  const requirements = ['Arial', 'Arial '].map((family, index) => ({
    key: `${family.toLowerCase()}|0|0`, family, bold: false, italic: false, characters: [index ? 'B' : 'A'],
  }));
  await source.prepareFontRequirements(requirements);
  const config = source.measurementConfigForRequirements(requirements);
  expect(config).toBeDefined();
  expect(config?.fontChains['arial|0|0']).toEqual(config?.fontChains['arial |0|0']);
  expect(await source.prepareFontRequirements(requirements)).toBe(false);
  expect(source.measurementConfigForRequirements(requirements)).toEqual(config);
});

test('compact requirement bitmaps retain candidate boundaries and aliases', async () => {
  const characters = [...'AİΩЯ‐₹↔∀■✓ﬀ\uf06f'];
  const bits = coverageBits(character => characters.includes(character));
  expect(charactersFromBits(bits).sort()).toEqual(characters.sort());
  expect(bits.length).toBeLessThanOrEqual(512);
  const load = mock(async () => regular);
  let next = 0;
  const source = createRustMeasureSource({ engine: { registerFont: () => next++, clearFonts() {} },
    bundled: { resolve: () => async () => hebrew, resolveBaseLatin: () => load } });
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false,
    characterBits: coverageBits(character => character === 'Ω') };
  await source.prepareFontRequirements([requirement]);
  expect(source.measurementConfigForRequirements([requirement])?.fontChains[requirement.key]).toEqual([0, 1]);
  expect(load).toHaveBeenCalledTimes(1);
});

for (const tier of ['embedded', 'configured', 'bundled', 'script'] as const) {
  test(`base Latin stays unloaded when the ${tier} tier covers the text`, async () => {
    const load = mock(async () => regular);
    let next = 0;
    const registry = new TextMeasureFontRegistry({ registerFont: () => next++ }, { bundled: {
      resolve: () => tier === 'bundled' ? async () => primary : undefined,
      resolveScriptFallback: () => tier === 'script' ? async () => primary : undefined,
      resolveBaseLatin: () => load,
    } });
    if (tier === 'embedded' || tier === 'configured') registry.setEmbeddedFaces([{
      family: 'Requested', weight: 'normal', style: 'normal', subsetted: true, data: primary,
    }]);
    const ids = await registry.getFontIdChain('Requested', false, false);
    ids.push(...await registry.getScriptFallbackIds(tier === 'script' ? ['hebrew'] : []));
    await registry.prepareBaseLatin('Requested', false, false, [...'Latin Ω'], ids);
    expect(load).not.toHaveBeenCalled();
    expect(registry.getCachedBaseLatinIds('Requested', false, false)).toEqual([]);
  });
}

test('uncovered text loads one normal face and appends it after script fallbacks', async () => {
  const load = mock(async () => regular);
  const registered: Uint8Array[] = [];
  const source = createRustMeasureSource({
    engine: { registerFont(bytes) { registered.push(bytes); return registered.length - 1; }, clearFonts() {} },
    bundled: {
      resolve: () => async () => hebrew,
      resolveScriptFallback: () => async () => hebrew.slice(0),
      resolveBaseLatin: () => load,
    },
  });
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false,
    scripts: ['hebrew' as const], characters: ['Ω'] };
  expect(fontCoverage(new Uint8Array(hebrew))?.('Ω')).toBe(false);
  await source.prepareFontRequirements([requirement]);
  expect(source.measurementConfigForRequirements([requirement])?.fontChains[requirement.key]).toEqual([0, 1, 2]);
  expect(load).toHaveBeenCalledTimes(1);
  expect(registered[2]).toEqual(new Uint8Array(regular));
  await source.prepareFontRequirements([requirement]);
  expect(load).toHaveBeenCalledTimes(1);
  expect(registered).toHaveLength(3);
});

test('a styled chain loads Regular directly for an uncovered Greek character', async () => {
  const requested: Array<[boolean, boolean]> = [];
  let next = 0;
  const registry = new TextMeasureFontRegistry({ registerFont: () => next++ }, { bundled: {
    resolve: () => async () => hebrew,
    resolveBaseLatin: (bold, italic) => {
      requested.push([bold, italic]);
      return async () => regular;
    },
  } });
  const ids = await registry.getFontIdChain('Requested', true, true);
  await registry.prepareBaseLatin('Requested', true, true, ['Ω'], ids);
  expect(requested).toEqual([[false, false]]);
  expect(registry.getCachedBaseLatinIds('Requested', true, true)).toEqual([1]);
});

test('an empty chain stays empty and never loads base Latin', async () => {
  const load = mock(async () => regular);
  const registerFont = mock(() => 0);
  const source = createRustMeasureSource({
    engine: { registerFont, clearFonts() {} },
    bundled: { resolve: () => undefined, resolveBaseLatin: () => load },
  });
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['A', 'Ω'] };
  await source.prepareFontRequirements([requirement]);
  expect(source.measurementConfigForRequirements([requirement])?.fontChains).toEqual({});
  expect(source.measurementConfigForRequirements([{ ...requirement, characters: ['A', 'Ω', 'B'] }])?.fontChains).toEqual({});
  expect(load).not.toHaveBeenCalled();
  expect(registerFont).not.toHaveBeenCalled();
});

test('an unavailable base asset preserves earlier chains and retries without breaking preparation', async () => {
  let unavailable = true;
  const load = mock(async () => {
    if (unavailable) throw new Error('offline');
    return regular;
  });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const provider: BundledFontProvider = { resolve: () => async () => hebrew, resolveBaseLatin: () => load };
  let next = 0;
  const source = createRustMeasureSource({
    engine: { registerFont: () => next++, clearFonts() {} }, bundled: provider,
  });
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['Ω'] };
  try {
    await expect(source.prepareFontRequirements([requirement])).resolves.toBe(true);
    expect(source.measurementConfigForRequirements([requirement])?.fontChains).toEqual({ [requirement.key]: [0] });
    unavailable = false;
    await source.prepareFontRequirements([requirement]);
    expect(source.measurementConfigForRequirements([requirement])?.fontChains).toEqual({ [requirement.key]: [0, 1] });
    expect(load).toHaveBeenCalledTimes(2);
  } finally { warn.mockRestore(); }
});

test('private-use bullet aliases load only for symbol families', async () => {
  const load = mock(async () => regular);
  let next = 0;
  const registry = new TextMeasureFontRegistry({ registerFont: () => next++ }, {
    bundled: { resolve: () => async () => hebrew, resolveBaseLatin: () => load },
  });
  const chain = async (family: string) => [...await registry.getFontIdChain(family, false, false)];
  for (const family of ['Custom Logo', 'Arial', 'Wingdings 4']) {
    await registry.prepareBaseLatin(family, false, false, ['\uf06f'], await chain(family));
    expect(registry.getCachedBaseLatinIds(family, false, false)).toEqual([]);
  }
  expect(load).not.toHaveBeenCalled();
  for (const family of ['Symbol', 'sYmBoL', 'Wingdings', 'WINGDINGS 2', 'wingdings 3']) {
    await registry.prepareBaseLatin(family, false, false, ['\uf06f'], await chain(family));
    expect(registry.getCachedBaseLatinIds(family, false, false)).toEqual([1]);
  }
  expect(load).toHaveBeenCalledTimes(1);
});

test('new text invalidates prepared coverage even when family and style stay the same', async () => {
  const load = mock(async () => regular);
  let next = 0;
  const source = createRustMeasureSource({
    engine: { registerFont: () => next++, clearFonts() {} },
    bundled: { resolve: () => async () => hebrew, resolveBaseLatin: () => load },
  });
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['א'] };
  await source.prepareFontRequirements([requirement, { ...requirement, characters: undefined }]);
  expect(source.measurementConfigForRequirements([requirement])?.fontChains[requirement.key]).toEqual([0]);
  expect(load).not.toHaveBeenCalled();
  const changed = { ...requirement, characters: ['א', 'Ω'] };
  expect(source.measurementConfigForRequirements([changed])).toBeUndefined();
  await source.prepareFontRequirements([changed]);
  expect(source.measurementConfigForRequirements([changed])?.fontChains[changed.key]).toEqual([0, 1]);
  expect(load).toHaveBeenCalledTimes(1);
});

for (const signature of [0x774f4646, 0x774f4632]) {
  test(`compressed app font ${signature.toString(16)} keeps its bytes and chain without the tier`, async () => {
    const data = new ArrayBuffer(16);
    new DataView(data).setUint32(0, signature);
    const load = mock(async () => regular);
    const registerFont = mock((_bytes: Uint8Array) => 0);
    const source = createRustMeasureSource({
      engine: { registerFont, clearFonts() {} },
      bundled: { resolve: () => undefined, resolveBaseLatin: () => load },
    });
    source.setEmbeddedFaces([], [{ family: 'Requested', data, weight: 'normal', style: 'normal', subsetted: false }]);
    const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['Ω'] };
    await source.prepareFontRequirements([requirement]);
    const config = source.measurementConfigForRequirements([requirement]);
    expect(config?.fontChains[requirement.key]).toEqual([0]);
    expect(source.measurementConfigForRequirements([{ ...requirement, characters: ['Ω', 'Ж'] }])).toEqual(config);
    expect(registerFont).toHaveBeenCalledTimes(1);
    expect(registerFont).toHaveBeenCalledWith(new Uint8Array(data));
    expect(fontCoverage(new Uint8Array(data))).toBeUndefined();
    expect(load).not.toHaveBeenCalled();
  });
}

test('an uncompressed app subset loads the tier for an uncovered Greek character', async () => {
  const load = mock(async () => regular);
  const registered: Uint8Array[] = [];
  const source = createRustMeasureSource({
    engine: { registerFont(bytes) { return registered.push(bytes) - 1; }, clearFonts() {} },
    bundled: { resolve: () => undefined, resolveBaseLatin: () => load },
  });
  source.setEmbeddedFaces([], [{ family: 'Requested', data: hebrew, weight: 'normal', style: 'normal', subsetted: true }]);
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['Ω'] };
  await source.prepareFontRequirements([requirement]);
  expect(source.measurementConfigForRequirements([requirement])?.fontChains[requirement.key]).toEqual([0, 1]);
  expect(registered).toEqual([new Uint8Array(hebrew), new Uint8Array(regular)]);
  expect(load).toHaveBeenCalledTimes(1);
});

test('unknown coverage anywhere in a chain removes a cached backstop', async () => {
  const data = new ArrayBuffer(16);
  new DataView(data).setUint32(0, 0x774f4632);
  const load = mock(async () => regular);
  let next = 0;
  const registry = new TextMeasureFontRegistry({ registerFont: () => next++ }, { bundled: {
    resolve: () => async () => hebrew,
    resolveScriptFallback: () => async () => data,
    resolveBaseLatin: () => load,
  } });
  const ids = await registry.getFontIdChain('Requested', false, false);
  await registry.prepareBaseLatin('Requested', false, false, ['Ω'], ids);
  expect(registry.getCachedBaseLatinIds('Requested', false, false, ['Ω'], ids)).toEqual([1]);
  const unknown = [...ids, ...await registry.getScriptFallbackIds(['hebrew'])];
  expect(registry.getCachedBaseLatinIds('Requested', false, false, ['Ω'], unknown)).toEqual([]);
  expect(await registry.prepareBaseLatin('Requested', false, false, ['Ω'], unknown)).toBe(true);
  expect(registry.getCachedBaseLatinIds('Requested', false, false, ['Ω', 'Ж'], unknown)).toEqual([]);
  expect(load).toHaveBeenCalledTimes(1);
});

test('uncompressed collections inspect the first face without decoding', () => {
  const bytes = new Uint8Array(primary.byteLength + 16);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x74746366);
  view.setUint32(4, 0x00010000);
  view.setUint32(8, 1);
  view.setUint32(12, 16);
  bytes.set(new Uint8Array(primary), 16);
  for (let index = 0; index < view.getUint16(20); index++) {
    const record = 28 + index * 16;
    view.setUint32(record + 8, view.getUint32(record + 8) + 16);
  }
  const covers = fontCoverage(bytes);
  expect(covers?.('A')).toBe(true);
  expect(covers?.('Ω')).toBe(true);
  expect(covers?.('日')).toBe(false);
});

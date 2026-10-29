import { afterEach, describe, expect, mock, test } from 'bun:test';
import { INTL_FONT_ASSET_URLS } from '../../fonts-intl/src';
import { version as fontsIntlVersion } from '../../fonts-intl/package.json';
import { version as fontsVersion } from '../package.json';
import intlPackage from '../../fonts-intl/package.json';
import cjkPackage from '../../fonts-cjk/package.json';
import { TextMeasureFontRegistry } from '../../docx/src/layout/measure/fontRegistry';
import {
  BUNDLED_FONTS,
  resolveBundledFamilyFace,
  resolveLastResortFace,
  resolveMetricCompatFace,
  resolveMetricCompatFamily,
  resolveScriptFallbackFace,
  type BundledFontScript,
} from './manifest';
import { SCRIPT_PATTERNS } from './scripts';

const families: readonly [string, BundledFontScript, boolean][] = [
  ['Noto Sans Thai', 'thai', true],
  ['Noto Sans Lao', 'lao', true],
  ['Noto Sans Khmer', 'khmer', true],
  ['Noto Sans Myanmar', 'myanmar', true],
  ['Noto Sans Devanagari', 'devanagari', true],
  ['Noto Sans Bengali', 'bengali', true],
  ['Noto Sans Gurmukhi', 'gurmukhi', true],
  ['Noto Sans Gujarati', 'gujarati', true],
  ['Noto Sans Oriya', 'oriya', true],
  ['Noto Sans Tamil', 'tamil', true],
  ['Noto Sans Telugu', 'telugu', true],
  ['Noto Sans Kannada', 'kannada', true],
  ['Noto Sans Malayalam', 'malayalam', true],
  ['Noto Sans Sinhala', 'sinhala', true],
  ['Noto Sans Armenian', 'armenian', true],
  ['Noto Sans Georgian', 'georgian', true],
  ['Noto Sans Ethiopic', 'ethiopic', true],
  ['Noto Serif Tibetan', 'tibetan', true],
  ['Noto Sans Mongolian', 'mongolian', false],
  ['Noto Sans Syriac', 'syriac', false],
  ['Noto Sans Thaana', 'thaana', true],
  ['Noto Sans Cherokee', 'cherokee', true],
  ['Noto Sans Lao Looped', 'lao', true],
  ['Noto Sans Tifinagh', 'tifinagh', false],
];
const intl = BUNDLED_FONTS.filter((face) => face.file.endsWith('.ttf.zst'));

describe('international font manifest', () => {
  test('ships the pack, licenses, and built entry points at the CJK package version', () => {
    expect(intlPackage.version).toBe(cjkPackage.version);
    expect(intlPackage.exports['.'].import).toBe('./dist/index.js');
    expect(intlPackage.exports['.'].types).toBe('./dist/index.d.ts');
    expect(intlPackage.files).toEqual(['dist', 'assets', 'LICENSE', 'LICENSES', 'README.md']);
  });

  test('contains each requested family and available style exactly once', () => {
    expect(intl).toHaveLength(45);
    expect(new Set(intl.map((face) => face.file)).size).toBe(intl.length);
    expect(new Set(BUNDLED_FONTS.map((face) => face.file)).size).toBe(BUNDLED_FONTS.length);
    expect(new Set(intl.map((face) => face.family)).size).toBe(24);
    for (const [family, script, bold] of families) {
      const faces = intl.filter((face) => face.family === family);
      expect(faces.map((face) => face.weight)).toEqual(bold ? [400, 700] : [400]);
      for (const face of faces) {
        expect(face.script).toBe(script);
        expect(face.metricCompatWith).toBeUndefined();
        expect(face.style).toBe('normal');
        expect(face.file).toBe(
          `${family.replace(/ /g, '')}-${face.weight === 700 ? 'Bold' : 'Regular'}.ttf.zst`,
        );
      }
    }
  });

  test('has exactly one literal add-on URL for every intl manifest row', async () => {
    const files = intl.map((face) => face.file).sort();
    expect(Object.keys(INTL_FONT_ASSET_URLS).sort()).toEqual(files);
    const source = await Bun.file(new URL('../../fonts-intl/src/index.ts', import.meta.url)).text();
    const rows = [...source.matchAll(/'([^']+)': \(\) => new URL\('\.\.\/assets\/([^']+)', import\.meta\.url\)/g)];
    expect(rows.map(([, file]) => file).sort()).toEqual(files);
    for (const [, file, target] of rows) {
      expect(file).toBe(target);
      expect(INTL_FONT_ASSET_URLS[file]().pathname).toEndWith(`/assets/${file}`);
    }
  });

  test('pins exactly the source TTF for each packed face', async () => {
    const checksums = await Bun.file(new URL('../../../scripts/fonts/intl-sources.sha256', import.meta.url)).text();
    const files = checksums.trim().split('\n').map((line) => line.trim().split(/\s+/)[1]);
    expect(files.sort()).toEqual(intl.map((face) => face.file.replace(/\.zst$/, '')).sort());
    expect(new Set(files).size).toBe(files.length);
  });

  for (const [script, character] of [
    ['thai', 'ก'],
    ['lao', 'ກ'],
    ['khmer', 'ក'],
    ['myanmar', 'က'],
    ['devanagari', 'क'],
    ['bengali', 'ক'],
    ['gurmukhi', 'ਕ'],
    ['gujarati', 'ક'],
    ['oriya', 'କ'],
    ['tamil', 'க'],
    ['telugu', 'క'],
    ['kannada', 'ಕ'],
    ['malayalam', 'ക'],
    ['sinhala', 'ක'],
    ['armenian', 'Ա'],
    ['georgian', 'ა'],
    ['ethiopic', 'ሀ'],
    ['tibetan', 'ཀ'],
    ['mongolian', 'ᠠ'],
    ['syriac', 'ܐ'],
    ['thaana', 'ހ'],
    ['cherokee', 'Ꭰ'],
    ['tifinagh', 'ⴰ'],
  ] as const) {
    test(`detects ${script} and resolves its fallback`, () => {
      const detected = SCRIPT_PATTERNS.filter(([, pattern]) => pattern.test(character));
      expect(detected.map(([name]) => name)).toEqual([script]);
      const family = families.find(([, bucket]) => bucket === script)!;
      for (const bold of [false, true]) {
        for (const italic of [false, true]) {
          const face = resolveScriptFallbackFace(script, bold, italic)!;
          expect(face.family).toBe(family[0]);
          expect(face.weight).toBe(bold && family[2] ? 700 : 400);
          expect(face.style).toBe('normal');
        }
      }
    });
  }

  for (const [family, officeNames] of [
    ['Noto Sans Thai', ['Leelawadee UI', 'Leelawadee', 'Angsana New', 'Browallia New', 'Cordia New', 'DilleniaUPC']],
    ['Noto Sans Lao', ['Lao UI', 'DokChampa']],
    ['Noto Sans Khmer', ['Khmer UI', 'DaunPenh', 'MoolBoran']],
    ['Noto Sans Myanmar', ['Myanmar Text']],
    ['Noto Sans Devanagari', ['Nirmala UI', 'Mangal', 'Aparajita', 'Kokila', 'Utsaah']],
    ['Noto Sans Bengali', ['Vrinda', 'Shonar Bangla']],
    ['Noto Sans Gurmukhi', ['Raavi']],
    ['Noto Sans Gujarati', ['Shruti']],
    ['Noto Sans Oriya', ['Kalinga']],
    ['Noto Sans Tamil', ['Latha', 'Vijaya']],
    ['Noto Sans Telugu', ['Gautami', 'Vani']],
    ['Noto Sans Kannada', ['Tunga']],
    ['Noto Sans Malayalam', ['Kartika']],
    ['Noto Sans Sinhala', ['Iskoola Pota']],
    ['Noto Sans Ethiopic', ['Nyala']],
    ['Noto Serif Tibetan', ['Microsoft Himalaya']],
    ['Noto Sans Mongolian', ['Mongolian Baiti']],
    ['Noto Sans Syriac', ['Estrangelo Edessa']],
    ['Noto Sans Thaana', ['MV Boli']],
    ['Noto Sans Cherokee', ['Plantagenet Cherokee']],
  ] as const) {
    test(`uses ${family} only through script fallback for Office names`, () => {
      const [, script, hasBold] = families.find(([name]) => name === family)!;
      for (const office of officeNames) {
        for (const bold of [false, true]) {
          expect(resolveMetricCompatFamily(office)).toBeUndefined();
          expect(resolveMetricCompatFace(` ${office.toUpperCase()} `, bold, false)).toBeUndefined();
          expect(resolveBundledFamilyFace(office, bold, false)).toBeUndefined();
          expect(resolveBundledFamilyFace(office, bold, true)).toBeUndefined();
          const face = resolveScriptFallbackFace(script, bold, false)!;
          expect(face.family).toBe(family);
          expect(face.weight).toBe(bold && hasBold ? 700 : 400);
        }
      }
    });
  }

  test('leaves multiscript and Latin Office families on their existing family paths', () => {
    for (const name of ['Sylfaen', 'Ebrima', 'Gadugi', 'Tahoma']) {
      expect(resolveMetricCompatFamily(name)).toBeUndefined();
      expect(resolveBundledFamilyFace(name, false, false)).toBeUndefined();
    }
    expect(resolveMetricCompatFace('Arial', false, false)?.family).toBe('Liberation Sans');
    expect(resolveScriptFallbackFace('lao', false, false)?.family).toBe('Noto Sans Lao');
    expect(resolveBundledFamilyFace('Noto Sans Lao Looped', true, false)?.file)
      .toBe('NotoSansLaoLooped-Bold.ttf.zst');
  });
});

for (const [family, character, script, fallback] of [
  ['Nirmala UI', 'क', 'devanagari', 'Noto Sans Devanagari'],
  ['Leelawadee UI', 'ก', 'thai', 'Noto Sans Thai'],
] as const) {
  test(`Latin-only ${family} keeps the main Liberation Sans chain and ${character} has a script fallback`, async () => {
    const { createFontProvider } = await import('./index');
    const provider = createFontProvider();
    const registered: Uint8Array[] = [];
    const registry = new TextMeasureFontRegistry(
      { registerFont: (bytes) => { registered.push(bytes); return registered.length; } },
      { bundled: provider },
    );
    expect(SCRIPT_PATTERNS.filter(([, pattern]) => pattern.test('Latin text 123'))).toEqual([]);
    expect(provider.resolve(family, false, false)).toBeUndefined();
    expect(await registry.getFontIdChain(family, false, false)).toEqual([1]);
    expect(registered).toEqual([new Uint8Array(await provider.resolve('Arial', false, false)!())]);
    expect(resolveLastResortFace(family, false, false, 'word').family).toBe('Liberation Sans');
    expect(SCRIPT_PATTERNS.filter(([, pattern]) => pattern.test(character)).map(([name]) => name)).toEqual([script]);
    expect(provider.resolveScriptFallback(script, false, false)).toBeDefined();
    expect(resolveScriptFallbackFace(script, false, false)?.family).toBe(fallback);
  });
}

test('script blocks include extensions and keep neighboring scripts distinct', () => {
  for (const [script, points] of [
    ['myanmar', [0x1000, 0x109f, 0xa9e0, 0xa9ff, 0xaa60, 0xaa7f, 0x116d0, 0x116ff]],
    ['devanagari', [0x0900, 0x097f, 0xa8e0, 0xa8ff, 0x11b00, 0x11b5f]],
    ['armenian', [0x0530, 0x058f, 0xfb13, 0xfb17]],
    ['georgian', [0x10a0, 0x10ff, 0x1c90, 0x1cbf, 0x2d00, 0x2d2f]],
    ['ethiopic', [0x1200, 0x137f, 0x1380, 0x139f, 0x2d80, 0x2ddf, 0xab00, 0xab2f, 0x1e7e0, 0x1e7ff]],
    ['tifinagh', [0x2d30, 0x2d7f]],
    ['cherokee', [0x13a0, 0x13ff, 0xab70, 0xabbf]],
    ['syriac', [0x0700, 0x074f, 0x0860, 0x086f]],
    ['thaana', [0x0780, 0x07bf]],
    ['sinhala', [0x0d80, 0x0dff, 0x111e0, 0x111ff]],
    ['tamil', [0x0b80, 0x0bff, 0x11fc0, 0x11fff]],
    ['mongolian', [0x1800, 0x18af, 0x11660, 0x1167f]],
    ['khmer', [0x1780, 0x17ff, 0x19e0, 0x19ff]],
  ] as const) {
    for (const point of points) {
      expect(SCRIPT_PATTERNS.filter(([, pattern]) => pattern.test(String.fromCodePoint(point)))
        .map(([name]) => name)).toEqual([script]);
    }
  }
  for (const text of ['Latin', 'Ελληνικά', 'Кириллица', 'ᚠ', '\u2d7f\u2d80']) {
    const scripts = SCRIPT_PATTERNS.filter(([, pattern]) => pattern.test(text)).map(([name]) => name);
    expect(scripts).toEqual(text === '\u2d7f\u2d80' ? ['ethiopic', 'tifinagh'] : []);
  }
});

const decodedRegular = new Uint8Array([0, 1, 0, 0, 10]);
const decodedBold = new Uint8Array([0, 1, 0, 0, 20]);
const decode = mock(async (
  input: ArrayBuffer | Uint8Array,
  dictionary?: ArrayBuffer | Uint8Array,
  options?: { codec?: string; wasmUrl?: URL },
): Promise<ArrayBuffer> => {
  expect(options?.codec).toBe('zstd');
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (bytes[0] === 1) {
    expect(dictionary).toBeUndefined();
    return decodedRegular.slice().buffer;
  }
  expect(bytes[0]).toBe(2);
  expect(new Uint8Array(dictionary as ArrayBuffer)).toEqual(decodedRegular);
  return decodedBold.slice().buffer;
});
mock.module('./decode', () => ({ decodeFontBytes: decode }));

const importIntl = mock(() => ({
  INTL_FONT_ASSET_URLS: Object.fromEntries(
    intl.map((face) => [face.file, () => new URL(`https://intl-addon.example/assets/${face.file}`)]),
  ),
}));
mock.module('@betteroffice/fonts-intl', () => ({
  then(resolve: (addon: ReturnType<typeof importIntl>) => void, reject: (error: unknown) => void) {
    Promise.resolve().then(importIntl).then(resolve, reject);
  },
}));

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
  decode.mockClear();
});

test('lazily imports intl and decodes Bold with a cached decoded Regular dictionary', async () => {
  const { createFontProvider, loadBundledFontBytes } = await import('./index');
  const regular = resolveScriptFallbackFace('thai', false, false)!;
  const bold = resolveScriptFallbackFace('thai', true, false)!;
  const sizes = [regular.byteLength, bold.byteLength];
  regular.byteLength = bold.byteLength = 5;
  const requested: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requested.push(String(input));
    return new Response(new Uint8Array([String(input).includes('-Bold.') ? 2 : 1]));
  }) as typeof fetch;
  try {
    const provider = createFontProvider();
    const load = provider.resolveScriptFallback('thai', true, false)!;
    expect(importIntl).not.toHaveBeenCalled();
    expect(requested).toEqual([]);
    const [first, second, normal] = await Promise.all([
      load(), load(), provider.resolveScriptFallback('thai', false, false)!(),
    ]);
    expect(importIntl).toHaveBeenCalledTimes(1);
    expect(decode).toHaveBeenCalledTimes(2);
    expect(requested.sort()).toEqual([
      `https://intl-addon.example/assets/${bold.file}`,
      `https://intl-addon.example/assets/${regular.file}`,
    ].sort());
    expect(new Uint8Array(first)).toEqual(decodedBold);
    expect(first).not.toBe(second);
    structuredClone(first, { transfer: [first] });
    structuredClone(normal, { transfer: [normal] });
    expect(new Uint8Array(await load())).toEqual(decodedBold);
    expect(new Uint8Array(await loadBundledFontBytes(regular))).toEqual(decodedRegular);
    expect(decode).toHaveBeenCalledTimes(2);
  } finally {
    [regular.byteLength, bold.byteLength] = sizes;
  }
});

test('retries failed decoding and isolates dictionaries by asset root', async () => {
  const { loadBundledFontBytes } = await import('./index');
  const regular = resolveScriptFallbackFace('tamil', false, false)!;
  const bold = resolveScriptFallbackFace('tamil', true, false)!;
  const sizes = [regular.byteLength, bold.byteLength];
  regular.byteLength = bold.byteLength = 5;
  let attempts = 0;
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    attempts++;
    return new Response(new Uint8Array([String(input).includes('-Bold.') ? 2 : 1]));
  }) as typeof fetch;
  try {
    decode.mockImplementationOnce(async () => { throw new Error('corrupt Regular'); });
    const options = { baseUrl: 'https://intl-retry.example/' };
    await expect(loadBundledFontBytes(bold, options)).rejects.toThrow('corrupt Regular');
    expect(new Uint8Array(await loadBundledFontBytes(bold, options))).toEqual(decodedBold);
    expect(attempts).toBe(4);
    const other = { baseUrl: 'https://intl-other.example/' };
    expect(new Uint8Array(await loadBundledFontBytes(bold, other))).toEqual(decodedBold);
    expect(attempts).toBe(6);
    const normal = await loadBundledFontBytes(regular, options);
    structuredClone(normal, { transfer: [normal] });
    expect(new Uint8Array(await loadBundledFontBytes(bold, options))).toEqual(decodedBold);
  } finally {
    [regular.byteLength, bold.byteLength] = sizes;
  }
});

test('CDN loads the pinned intl pack and custom roots without importing the add-on', async () => {
  const { createFontProvider } = await import('./cdn');
  const regular = resolveScriptFallbackFace('bengali', false, false)!;
  const bold = resolveScriptFallbackFace('bengali', true, false)!;
  const sizes = [regular.byteLength, bold.byteLength];
  regular.byteLength = bold.byteLength = 5;
  const requested: string[] = [];
  const imports = importIntl.mock.calls.length;
  globalThis.fetch = (async (input: RequestInfo | URL, options?: RequestInit) => {
    expect(options?.credentials).toBe('omit');
    expect(options?.referrerPolicy).toBe('no-referrer');
    expect(options?.signal).toBeInstanceOf(AbortSignal);
    requested.push(String(input));
    return new Response(new Uint8Array([String(input).includes('-Bold.') ? 2 : 1]));
  }) as typeof fetch;
  try {
    const provider = createFontProvider();
    expect(requested).toEqual([]);
    const bytes = await provider.resolveScriptFallback('bengali', true, false)!();
    expect(new Uint8Array(bytes)).toEqual(decodedBold);
    const base = `https://cdn.jsdelivr.net/npm/@betteroffice/fonts-intl@${fontsIntlVersion}/assets/`;
    expect(requested.sort()).toEqual([`${base}${bold.file}`, `${base}${regular.file}`].sort());
    const custom = createFontProvider({ baseUrl: 'https://intl-custom-cdn.example/fonts' });
    expect(new Uint8Array(await custom.resolveScriptFallback('bengali', true, false)!()))
      .toEqual(decodedBold);
    expect(requested.slice(2).sort()).toEqual([
      `https://intl-custom-cdn.example/fonts/${bold.file}`,
      `https://intl-custom-cdn.example/fonts/${regular.file}`,
    ].sort());
    expect(decode.mock.calls.map(([, , options]) => options?.wasmUrl?.href)).toEqual(
      Array(4).fill(`https://cdn.jsdelivr.net/npm/@betteroffice/fonts@${fontsVersion}/dist/decoder/fontpack.wasm`),
    );
    expect(importIntl).toHaveBeenCalledTimes(imports);
  } finally {
    [regular.byteLength, bold.byteLength] = sizes;
  }
});

test('rejects HTTP failures and decoded size mismatches without caching failed faces', async () => {
  const { loadBundledFontBytes } = await import('./index');
  const regular = resolveScriptFallbackFace('khmer', false, false)!;
  const size = regular.byteLength;
  regular.byteLength = 5;
  let attempts = 0;
  globalThis.fetch = (async () => {
    attempts++;
    return attempts === 1
      ? new Response('unavailable', { status: 503 })
      : new Response(new Uint8Array([1]));
  }) as unknown as typeof fetch;
  try {
    const options = { baseUrl: 'https://intl-invalid-size.example/' };
    await expect(loadBundledFontBytes(regular, options)).rejects.toThrow('HTTP 503');
    decode.mockImplementationOnce(async () => new ArrayBuffer(4));
    await expect(loadBundledFontBytes(regular, options)).rejects.toThrow('expected 5');
    expect(new Uint8Array(await loadBundledFontBytes(regular, options))).toEqual(decodedRegular);
    expect(attempts).toBe(3);
  } finally {
    regular.byteLength = size;
  }
});

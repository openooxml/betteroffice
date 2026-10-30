import { fontAssetResponse } from './test-assets';
import { describe, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { existsSync } from 'node:fs';

import latinManifest from '../../fonts-latin/assets/manifest.json';
import manifest from '../package.json' with { type: 'json' };
import {
  BUNDLED_FONTS,
  createFontProvider,
  loadBundledFontBytes,
  maxDigitWidth,
  resolveBundledFamilyFace,
  resolveDigitAdvance,
  resolveLastResortFace,
  resolveLightVariantFace,
  resolveGlyphFallbackFace,
  resolveMetricCompatFace,
  resolveMetricCompatFamily,
  resolveScriptFallbackFace,
} from './index';

const SFNT_TRUETYPE = 0x00010000;
const SFNT_CFF = 0x4f54544f;

function expectLatinSfnt(bytes: ArrayBuffer, file: string, family: string) {
  const entry = latinManifest.faces.find((face) => face.file === file)!;
  expect(entry.family).toBe(family);
  expect(bytes.byteLength).toBe(entry.byteLength);
  expect(createHash('sha256').update(new Uint8Array(bytes)).digest('hex')).toBe(entry.sha256);
  const view = new DataView(bytes);
  expect(view.getUint32(0)).toBe(SFNT_TRUETYPE);
  const tables = Array.from({ length: view.getUint16(4) }, (_, index) => 12 + index * 16);
  const nameTable = tables.find((offset) => view.getUint32(offset) === 0x6e616d65);
  expect(nameTable).toBeDefined();
  const offset = view.getUint32(nameTable! + 8);
  const strings = offset + view.getUint16(offset + 4);
  const families: string[] = [];
  for (let i = 0; i < view.getUint16(offset + 2); i++) {
    const record = offset + 6 + i * 12;
    const platform = view.getUint16(record);
    if ((platform === 0 || platform === 3) && view.getUint16(record + 6) === 1) {
      const length = view.getUint16(record + 8);
      const start = strings + view.getUint16(record + 10);
      families.push(new TextDecoder('utf-16be').decode(new Uint8Array(bytes, start, length)));
    }
  }
  expect(families).toContain(family);
}

describe('resolution', () => {
  test('maps the MS core fonts to their metric-compatible substitutes', () => {
    expect(resolveMetricCompatFamily('Calibri')).toBe('Carlito');
    expect(resolveMetricCompatFamily('  CAMBRIA ')).toBe('Caladea');
    expect(resolveMetricCompatFamily('Helvetica')).toBe('Liberation Sans');
    expect(resolveMetricCompatFamily('Times')).toBe('Liberation Serif');
    for (const family of ['Tms Rmn', 'Times Roman', 'TimesNewRomanPSMT', 'TimesNewRoman']) {
      expect(resolveMetricCompatFamily(family)).toBe('Liberation Serif');
    }
    for (const family of ['Helv', 'Arial MT', 'ArialMT']) {
      expect(resolveMetricCompatFamily(family)).toBe('Liberation Sans');
    }
    expect(resolveMetricCompatFamily('Nonesuch')).toBeUndefined();
  });

  test('falls back to a family regular when the exact style is not vendored', () => {
    expect(resolveMetricCompatFace('Calibri', true, true)?.file).toBe('Carlito-BoldItalic.ttf.zst');
    expect(resolveMetricCompatFace('SimSun', true, false)?.file).toBe('NotoSerifSC-Regular.otf.zst');
  });

  test('MS Japanese spellings resolve to the bundled Japanese face in every style', () => {
    for (const family of [
      'MS Mincho', 'ＭＳ 明朝', 'MS Gothic', 'ＭＳ ゴシック',
      'MS PMincho', 'ＭＳ Ｐ明朝', 'MS PGothic', 'ＭＳ Ｐゴシック', 'MS UI Gothic',
    ]) {
      expect(resolveMetricCompatFamily(` ${family.toUpperCase()} `)).toBe('Noto Sans JP');
      for (const bold of [false, true]) {
        for (const italic of [false, true]) {
          expect(resolveMetricCompatFace(family, bold, italic)?.file).toBe('NotoSansJP-Regular.otf.zst');
        }
      }
    }
  });

  for (const [family, regular, boldFile] of [
    ['Segoe UI', 'Regular', 'Bold'],
    ['Segoe UI Light', 'Light', 'Light'],
    ['Segoe UI Semilight', 'Semilight', 'Semilight'],
    ['Segoe UI Semibold', 'Semibold', 'Bold'],
  ]) {
    test(`${family} resolves to Selawik for every upright style`, () => {
      expect(resolveMetricCompatFamily(` ${family.toUpperCase()} `)).toBe('Selawik');
      for (const bold of [false, true]) {
        const file = `Selawik-${bold ? boldFile : regular}.ttf.zst`;
        const face = resolveMetricCompatFace(family, bold, false);
        expect(face?.file).toBe(file);
        expect(face?.style).toBe('normal');
        expect(resolveBundledFamilyFace(family, bold, false)?.file).toBe(file);
        for (const office of ['word', 'powerpoint'] as const) {
          const coverageBold = / (Light|Semilight)$/.test(family) ? false : bold;
          const base = resolveMetricCompatFace(office === 'word' ? 'Arial' : 'Calibri', coverageBold, false)!;
          expect(resolveLastResortFace(family, bold, false, office).file).toBe(base.file);
          expect(resolveGlyphFallbackFace(family, bold, false, office)?.file).toBe(base.file);
        }
      }
    });
  }

  test('every metric-compatible request has a different last-resort file', () => {
    for (const face of BUNDLED_FONTS) {
      if (!face.metricCompatWith) continue;
      for (const bold of [false, true]) {
        for (const italic of [false, true]) {
          const primary = resolveMetricCompatFace(face.metricCompatWith, bold, italic);
          if (!primary) {
            expect([face.family, italic]).toEqual(['Selawik', true]);
            continue;
          }
          for (const office of ['word', 'powerpoint'] as const) {
            expect(resolveLastResortFace(face.metricCompatWith, bold, italic, office).file).not.toBe(primary.file);
          }
        }
      }
    }
  });

  test('Google Sans retains the Carlito fallback', () => {
    expect(resolveMetricCompatFamily('Google Sans')).toBeUndefined();
    expect(resolveLastResortFace('Google Sans', false, false).file).toBe('Carlito-Regular.ttf.zst');
    expect(resolveLastResortFace('Google Sans', true, true).file).toBe('Carlito-BoldItalic.ttf.zst');
  });

  test('Segoe UI Black follows the heavy family path', () => {
    expect(resolveMetricCompatFamily('Segoe UI Black')).toBeUndefined();
    expect(resolveLastResortFace('Segoe UI Black', false, false).file).toBe('Selawik-Bold.ttf.zst');
  });

  test('last resort always returns a face, serif-aware', () => {
    expect(resolveLastResortFace('Totally Unknown', false, false).family).toBe('Carlito');
    expect(resolveLastResortFace('Garamond', false, false).family).toBe('Liberation Serif');
    expect(resolveLastResortFace('Unknown', true, true).file).toBe('Carlito-BoldItalic.ttf.zst');
    expect(resolveLastResortFace('Calibri Light', false, false).file).toBe('Carlito-Regular.ttf.zst');
    expect(resolveLastResortFace(' CALIBRI LIGHT ', true, true).file).toBe('Carlito-Italic.ttf.zst');
    expect(resolveMetricCompatFamily('Calibri Light')).toBeUndefined();
  });

  test('light variants resolve only for families we bundle', () => {
    expect(resolveLightVariantFace('Calibri Light', false, false)?.file).toBe('Carlito-Regular.ttf.zst');
    expect(resolveLightVariantFace(' calibri light ', true, true)?.file).toBe('Carlito-Italic.ttf.zst');
    expect(resolveLightVariantFace('Segoe UI Semilight', false, false)?.family).toBe('Selawik');
    expect(resolveLightVariantFace('Calibri', false, false)).toBeUndefined();
    expect(resolveLightVariantFace('BIZ UD明朝 Medium', false, false)).toBeUndefined();
    expect(resolveLightVariantFace('Unknown Light', false, false)).toBeUndefined();
    expect(createFontProvider().resolveLightVariant?.('Calibri Light', true, false)).toBeFunction();
    expect(createFontProvider().resolveLightVariant?.('Wingdings', false, false)).toBeUndefined();
  });

  test('Word substitutes Arial for an unknown sans family, PowerPoint Calibri', () => {
    expect(resolveLastResortFace('Lato', false, false, 'word').file).toBe('LiberationSans-Regular.ttf.zst');
    expect(resolveLastResortFace('Lato', true, false, 'powerpoint').file).toBe('Carlito-Bold.ttf.zst');
    expect(resolveLastResortFace('Garamond', false, false, 'word').family).toBe('Liberation Serif');
    expect(resolveLastResortFace('Consolas', false, false, 'word').file).toBe('LiberationMono-Regular.ttf.zst');
  });

  for (const office of ['word', 'powerpoint'] as const) {
    test(`${office} keeps sans serif family names on the sans base`, () => {
      const sans = office === 'word' ? 'Arial' : 'Calibri';
      for (const family of [
        'Microsoft Sans Serif',
        'MS Sans Serif',
        'MS Reference Sans Serif',
        'Unknown Sans-Serif',
        'Unknown SansSerif',
        'UNKNOWN SANS SERIF',
      ]) {
        for (const [bold, italic] of [[false, false], [true, false], [false, true], [true, true]]) {
          expect(resolveLastResortFace(family, bold, italic, office)).toEqual(
            resolveMetricCompatFace(sans, bold, italic)!
          );
        }
      }
      for (const family of ['Garamond', 'Book Antiqua']) {
        expect(resolveLastResortFace(family, false, false, office).family).toBe('Liberation Serif');
      }
    });
  }

  test('a typewriter or old-style name lands on a face of its own kind', () => {
    expect(resolveLastResortFace('Consolas', false, false).file).toBe('LiberationMono-Regular.ttf.zst');
    expect(resolveLastResortFace('Lucida Console', false, false).file).toBe(
      'LiberationMono-Regular.ttf.zst'
    );
    expect(resolveLastResortFace('Book Antiqua', false, false).file).toBe(
      'LiberationSerif-Regular.ttf.zst'
    );
    expect(resolveLastResortFace('Bookman Old Style', true, false).file).toBe(
      'LiberationSerif-Bold.ttf.zst'
    );
    expect(resolveLastResortFace('Calisto MT', false, false).file).toBe(
      'LiberationSerif-Regular.ttf.zst'
    );
    expect(resolveLastResortFace('Monotype Corsiva', false, false).file).toBe(
      'Carlito-Regular.ttf.zst'
    );
  });

  test('a family with no clone takes the bundled face closest in width', () => {
    expect(resolveLastResortFace('Verdana', false, false).file).toBe('Montserrat-Regular.ttf.zst');
    expect(resolveLastResortFace('Verdana', true, false).file).toBe('Montserrat-Bold.ttf.zst');
    expect(resolveLastResortFace('Trebuchet MS', false, false).file).toBe('Carlito-Regular.ttf.zst');
    expect(resolveLastResortFace('Gill Sans MT', false, false).file).toBe('Carlito-Regular.ttf.zst');
  });

  test('a width-matched substitute keeps the digit advance of the family it draws', () => {
    expect(resolveDigitAdvance('Verdana')).toBe(0.636);
    expect(resolveDigitAdvance(' VERDANA ')).toBe(0.636);
    expect(resolveDigitAdvance('Calibri')).toBeUndefined();
    expect(resolveDigitAdvance('Trebuchet MS')).toBeUndefined();
  });

  test('the max digit width counts a substituted family in its own digits', () => {
    const montserrat: Record<string, number> = {
      0: 0.662, 1: 0.361, 2: 0.568, 3: 0.564, 4: 0.661,
      5: 0.566, 6: 0.609, 7: 0.589, 8: 0.638, 9: 0.609,
    };
    const drawn = (em: Record<string, number>, sizePx: number) => (digit: string) =>
      em[digit] * sizePx;
    expect(maxDigitWidth('Verdana', 10, drawn(montserrat, 10))).toBe(6);
    expect(maxDigitWidth('Verdana', 40 / 3, drawn(montserrat, 40 / 3))).toBe(8);
  });

  test('the max digit width respects a bold Normal font', () => {
    expect(maxDigitWidth('Verdana', 10, () => 8, true, false)).toBe(7);
    expect(resolveDigitAdvance('Verdana', true, false)).toBe(0.711);
  });

  test('the digit advance respects italic and bold italic styles', () => {
    expect(maxDigitWidth('Verdana', 10, () => 8, false, true)).toBe(6);
    expect(resolveDigitAdvance('Verdana', false, true)).toBe(0.636);
    expect(maxDigitWidth('Verdana', 10, () => 8, true, true)).toBe(7);
    expect(resolveDigitAdvance(' VERDANA ', true, true)).toBe(0.711);
  });

  test('unrecorded digit advances measure the widest digit', () => {
    const measure = (digit: string) => digit === '8' ? 7.5 : 6.4;
    expect(maxDigitWidth('Calibri', 10, measure)).toBe(8);
    expect(maxDigitWidth('Calibri', 10, measure, true, true)).toBe(8);
    expect(resolveDigitAdvance('Calibri', true, true)).toBeUndefined();
  });

  test('unrecorded digit advances round half pixels up', () => {
    const measure = (digit: string) => digit === '8' ? 6.5 : 6.4;
    expect(maxDigitWidth('Calibri', 10, measure)).toBe(7);
  });

  for (const family of ['constructor', '__proto__', 'toString', 'hasOwnProperty']) {
    test(`${family} has no recorded digit advance`, () => {
      for (const bold of [false, true]) {
        for (const italic of [false, true]) {
          expect(resolveDigitAdvance(family, bold, italic)).toBeUndefined();
        }
      }
    });

    test(`${family} measures every digit for its max width`, () => {
      for (const bold of [false, true]) {
        for (const italic of [false, true]) {
          const measured: string[] = [];
          const measure = (digit: string) => {
            measured.push(digit);
            return digit === '8' ? 7.5 : 6.4;
          };
          expect(maxDigitWidth(family, 10, measure, bold, italic)).toBe(8);
          expect(measured.join('')).toBe('0123456789');
        }
      }
    });

    test(`${family} resolves to the unknown-family fallback`, () => {
      for (const office of ['powerpoint', 'word'] as const) {
        for (const bold of [false, true]) {
          for (const italic of [false, true]) {
            expect(resolveLastResortFace(family, bold, italic, office)).toBe(
              resolveLastResortFace('Unknown Family', bold, italic, office),
            );
          }
        }
      }
    });
  }

  test('a trailing light weight on a known family outranks the bold flag', () => {
    expect(resolveLastResortFace('Calibri Light', true, false).file).toBe('Carlito-Regular.ttf.zst');
    expect(resolveLastResortFace('Inter Light', true, false).file).toBe('Inter-Regular.ttf.zst');
    expect(resolveLastResortFace('Archivo Black', true, false).file).toBe('Carlito-Bold.ttf.zst');
    expect(resolveLastResortFace('Lato Light', true, false).file).toBe('Carlito-Bold.ttf.zst');
    expect(resolveLastResortFace('Blackadder ITC', true, false).file).toBe('Carlito-Bold.ttf.zst');
  });

  test('metric-compatible clones resolve to the family they clone', () => {
    expect(resolveMetricCompatFamily('Arimo')).toBe('Liberation Sans');
    expect(resolveMetricCompatFamily('Tinos')).toBe('Liberation Serif');
    expect(resolveMetricCompatFamily('Carlito')).toBe('Carlito');
    expect(resolveMetricCompatFamily('Caladea')).toBe('Caladea');
  });

  test('script fallbacks prefer the sans face of the bucket', () => {
    expect(resolveScriptFallbackFace('cjk-sc', false, false)?.family).toBe('Noto Sans SC');
    expect(resolveScriptFallbackFace('arabic', false, false)?.family).toBe('Noto Sans Arabic');
    expect(resolveScriptFallbackFace('hebrew', true, false)?.file).toBe('NotoSansHebrew-Bold.ttf.zst');
  });

  test('known heavy family names select the bold face and preserve italics', () => {
    for (const family of ['Arial Black', ' ARIAL-BLACK ', 'Arial Heavy', 'Arial ExtraBold']) {
      expect(resolveLastResortFace(family, false, false).file).toBe('LiberationSans-Bold.ttf.zst');
      expect(resolveLastResortFace(family, false, true).file).toBe('LiberationSans-BoldItalic.ttf.zst');
    }
    expect(resolveLastResortFace('Calibri Heavy', false, false).file).toBe('Carlito-Bold.ttf.zst');
    expect(resolveLastResortFace('Calibri Heavy', false, true).file).toBe('Carlito-BoldItalic.ttf.zst');
  });

  test('unknown heavy family names keep the normal fallback weight', () => {
    expect(resolveLastResortFace('Archivo Black', false, false).file).toBe('Carlito-Regular.ttf.zst');
    expect(resolveLastResortFace('Archivo Black', false, true).file).toBe('Carlito-Italic.ttf.zst');
    expect(resolveLastResortFace('Archivo Black', true, false).file).toBe('Carlito-Bold.ttf.zst');
  });
});

describe('bundled family names', () => {
  test('resolve to their own faces, not only as Word families', () => {
    expect(resolveMetricCompatFace('Gelasio', false, false)).toBeUndefined();
    expect(resolveBundledFamilyFace('Gelasio', true, true)?.file).toBe('Gelasio-BoldItalic.ttf.zst');
    expect(resolveBundledFamilyFace(' source sans 3 ', false, true)?.file).toBe(
      'SourceSans3-Italic.ttf.zst'
    );
    expect(resolveBundledFamilyFace('Comic Relief', true, false)?.file).toBe('ComicRelief-Bold.ttf.zst');
    expect(resolveBundledFamilyFace('Noto Serif SC', false, false)?.file).toBe(
      'NotoSerifSC-Regular.otf.zst'
    );
  });

  test('never stand in another style, so the browser can synthesize it', () => {
    expect(resolveBundledFamilyFace('Oswald', false, true)).toBeUndefined();
    expect(resolveBundledFamilyFace('Noto Serif SC', true, false)).toBeUndefined();
    expect(resolveBundledFamilyFace('SimSun', true, false)).toBeUndefined();
    expect(resolveMetricCompatFace('SimSun', true, false)?.file).toBe('NotoSerifSC-Regular.otf.zst');
  });

  test('cover the serif Noto families the package does not vendor with the sans face', () => {
    expect(resolveBundledFamilyFace('Noto Serif TC', false, false)?.family).toBe('Noto Sans TC');
    expect(resolveBundledFamilyFace('Noto Serif JP', false, false)?.family).toBe('Noto Sans JP');
    expect(resolveBundledFamilyFace('Noto Serif KR', false, false)?.family).toBe('Noto Sans KR');
  });

  test('fall back to Word-family resolution and miss unknown names', () => {
    expect(resolveBundledFamilyFace('Calibri', true, true)?.file).toBe('Carlito-BoldItalic.ttf.zst');
    expect(resolveBundledFamilyFace('Tinos', false, false)?.family).toBe('Liberation Serif');
    expect(resolveBundledFamilyFace('Nonesuch', false, false)).toBeUndefined();
  });

  test('are served by the provider', async () => {
    const bytes = await createFontProvider().resolveFamily('Gelasio', false, false)!();
    expectLatinSfnt(bytes, 'Gelasio-Regular.ttf.zst', 'Gelasio');
    expect(createFontProvider().resolve('Gelasio', false, false)).toBeUndefined();
  });
});

describe('loading', () => {
  test('createFontProvider serves real sfnt bytes', async () => {
    const provider = createFontProvider();
    const bytes = await provider.resolve('Calibri', false, false)!();
    expectLatinSfnt(bytes, 'Carlito-Regular.ttf.zst', 'Carlito');
  });

  test('Selawik providers keep byte loads lazy and supply Office coverage', async () => {
    const { fontProvider } = await import('./provider');
    const loaded: string[] = [];
    const provider = fontProvider(async (face) => {
      loaded.push(face.file);
      return new ArrayBuffer(face.byteLength);
    });
    const regular = provider.resolve('Segoe UI', false, false)!;
    expect(regular.measureAs).toBe('Segoe UI');
    expect(provider.resolve('Calibri', false, false)!.measureAs).toBeUndefined();
    expect(provider.resolveLastResort('Arial', false, false, 'word').onlyIfPrimaryFails).toBe(true);
    expect(provider.resolveLastResort('Segoe UI', false, false, 'word').onlyIfPrimaryFails).toBeUndefined();
    expect(provider.resolveLastResort('Google Sans', false, false).onlyIfPrimaryFails).toBeUndefined();
    const light = provider.resolveFamily('Segoe UI Light', true, false)!;
    const fallback = provider.resolveGlyphFallback!('Segoe UI', false, false)!;
    expect(loaded).toEqual([]);
    await provider.resolveLastResort('Google Sans', false, false)();
    expect(loaded).toEqual(['Carlito-Regular.ttf.zst']);
    expect(provider.resolveGlyphFallback!('Google Sans', false, false)).toBeUndefined();
    await regular();
    await light();
    await fallback();
    expect(loaded).toEqual([
      'Carlito-Regular.ttf.zst', 'Selawik-Regular.ttf.zst', 'Selawik-Light.ttf.zst', 'Carlito-Regular.ttf.zst',
    ]);
    for (const office of ['powerpoint', 'word'] as const) {
      for (const family of ['Segoe UI', 'Segoe UI Light', 'Segoe UI Semilight', 'Segoe UI Semibold', 'Segoe UI Black']) {
        for (const bold of [false, true]) {
          expect(provider.resolve(family, bold, false)?.measureAs).toBe(
            family === 'Segoe UI Black' ? undefined : resolveMetricCompatFace(family, bold, false)!.metricCompatWith,
          );
          const coverageBold = / (Light|Semilight)$/.test(family) ? false : bold;
          const expected = resolveMetricCompatFace(office === 'word' ? 'Arial' : 'Calibri', coverageBold, false)!;
          expect(resolveGlyphFallbackFace(family, bold, false, office)).toBe(expected);
          await provider.resolveGlyphFallback!(family, bold, false, office)!();
          expect(loaded[loaded.length - 1]).toBe(expected.file);
          expect(provider.resolve(family, bold, true)).toBeUndefined();
          expect(resolveGlyphFallbackFace(family, bold, true, office)).toBeUndefined();
        }
      }
    }
  });

  test('italic Segoe UI draws a true italic face while upright Segoe UI draws Selawik', async () => {
    const provider = createFontProvider();
    const italicAngle = (bytes: ArrayBuffer) => {
      const view = new DataView(bytes);
      for (let i = 0; i < view.getUint16(4); i++) {
        const entry = 12 + i * 16;
        if (view.getUint32(entry) === 0x706f7374) return view.getInt32(view.getUint32(entry + 8) + 4) / 65536;
      }
      throw new Error('no post table');
    };
    for (const office of ['powerpoint', 'word'] as const) {
      for (const family of ['Segoe UI', 'Segoe UI Light', 'Segoe UI Semilight', 'Segoe UI Semibold']) {
        for (const bold of [false, true]) {
          expect(resolveMetricCompatFace(family, bold, false)?.family).toBe('Selawik');
          expect(resolveMetricCompatFace(family, bold, true)).toBeUndefined();
          expect(resolveBundledFamilyFace(family, bold, true)).toBeUndefined();
          const face = resolveLastResortFace(family, bold, true, office);
          expect(face.family).not.toBe('Selawik');
          expect(face.style).toBe('italic');
          expect(italicAngle(await provider.resolveLastResort(family, bold, true, office)())).toBeLessThan(0);
        }
      }
    }
  });

  test('the CJK add-on resolves through the optional dynamic import', async () => {
    const provider = createFontProvider();
    const bytes = await provider.resolveScriptFallback('cjk-sc', false, false)!();
    expect(new DataView(bytes).getUint32(0)).toBe(SFNT_CFF);
    const { script: _script, ...withoutScript } = resolveScriptFallbackFace('cjk-sc', false, false)!;
    expect(new DataView(await loadBundledFontBytes(withoutScript)).getUint32(0)).toBe(SFNT_CFF);
  });

  test('caches per face, handing out consumer-owned buffers', async () => {
    const face = resolveMetricCompatFace('Cambria', false, false)!;
    const [a, b] = await Promise.all([loadBundledFontBytes(face), loadBundledFontBytes(face)]);
    expect(a).not.toBe(b);
    expect(a).toEqual(b);
  });

  test('pins a relative provider base to the route where it is created', async () => {
    const originalLocation = Object.getOwnPropertyDescriptor(globalThis, 'location');
    const realFetch = globalThis.fetch;
    const face = resolveMetricCompatFace('Calibri', false, false)!;
    let requested = '';
    Object.defineProperty(globalThis, 'location', {
      configurable: true,
      value: new URL('https://example.test/docs/team/report/'),
    });
    try {
      globalThis.fetch = (async (input: RequestInfo | URL) => {
        if (!String(input).endsWith('.wasm')) requested = String(input);
        return fontAssetResponse(input);
      }) as unknown as typeof fetch;
      const provider = createFontProvider({ baseUrl: 'fonts/' });
      Object.defineProperty(globalThis, 'location', {
        configurable: true,
        value: new URL('https://example.test/'),
      });

      const bytes = await provider.resolve('Calibri', false, false)!();

      expect(requested).toBe(`https://example.test/docs/team/report/fonts/${face.file}`);
      expect(bytes.byteLength).toBe(face.byteLength);
    } finally {
      globalThis.fetch = realFetch;
      if (originalLocation) Object.defineProperty(globalThis, 'location', originalLocation);
      else delete (globalThis as { location?: Location }).location;
    }
  });

  test('explains that server-side provider bases must be absolute', () => {
    expect(() => createFontProvider({ baseUrl: 'fonts/' })).toThrow(
      'baseUrl must be absolute when no browser location exists'
    );
  });

  test('rejects truncated assets and retries the evicted load', async () => {
    const realFetch = globalThis.fetch;
    const face = resolveMetricCompatFace('Calibri', false, false)!;
    const intact = await Bun.file(new URL(`../../fonts-latin/assets/${face.file}`, import.meta.url)).arrayBuffer();
    let attempts = 0;
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      if (String(input).endsWith('.wasm')) return fontAssetResponse(input);
      attempts++;
      return new Response(
        attempts === 1 ? intact.slice(0, Math.floor(intact.byteLength / 2)) : intact
      );
    }) as unknown as typeof fetch;
    try {
      const options = { baseUrl: 'https://length-check.example/assets/' };
      await expect(loadBundledFontBytes(face, options)).rejects.toThrow(
        `expected ${intact.byteLength}`
      );
      expect((await loadBundledFontBytes(face, options)).byteLength).toBe(face.byteLength);
      expect(attempts).toBe(2);
    } finally {
      globalThis.fetch = realFetch;
    }
  });

  // Bun supports file: fetch, so throwing here proves the Node filesystem path is used.
  test('reads file: assets from disk without fetch', async () => {
    const realFetch = globalThis.fetch;
    globalThis.fetch = (() => {
      throw new Error('fetch must not be used for file: assets');
    }) as unknown as typeof fetch;
    try {
      const face = resolveMetricCompatFace('Courier New', false, false)!;
      const bytes = await loadBundledFontBytes(face);
      expect(bytes.byteLength).toBe(face.byteLength);
    } finally {
      globalThis.fetch = realFetch;
    }
  });
});

// Plain Node refuses to type-strip TypeScript under node_modules.
describe('published shape', () => {
  test('entry points are built JavaScript, never raw TypeScript', () => {
    for (const entry of [manifest.main, manifest.module, manifest.exports['.'].import]) {
      expect(entry).toMatch(/^\.\/dist\/.*\.js$/);
    }
    expect(manifest.types).toMatch(/^\.\/dist\/.*\.d\.ts$/);
    expect(manifest.files).toContain('dist');
    expect(manifest.files).not.toContain('src');
  });

  test('ships a licence text for every family whose binaries it carries', () => {
    for (const licence of [
      'OFL-Carlito.txt',
      'OFL-Selawik.txt',
      'OFL-Caladea.txt',
      'OFL-Liberation.txt',
      'OFL-NotoArabic.txt',
      'OFL-NotoSansHebrew.txt',
    ]) {
      expect(existsSync(new URL(`../LICENSES/${licence}`, import.meta.url))).toBe(true);
    }
  });

  test('carries no CJK binary — those ship in @betteroffice/fonts-cjk', () => {
    const cjk = BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-'));
    expect(cjk.length).toBeGreaterThan(0);
    for (const face of cjk) {
      expect(existsSync(new URL(`../assets/${face.file}`, import.meta.url))).toBe(false);
    }
  });

  test('carries no intl binary from the optional add-on', () => {
    const intl = BUNDLED_FONTS.filter((face) => face.file.endsWith('.ttf.zst'));
    expect(intl.length).toBeGreaterThan(0);
    for (const face of intl) {
      expect(existsSync(new URL(`../assets/${face.file}`, import.meta.url))).toBe(false);
    }
  });
});

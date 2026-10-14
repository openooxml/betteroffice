import { describe, expect, test } from 'bun:test';
import { existsSync, readdirSync } from 'node:fs';

import manifest from '../package.json' with { type: 'json' };
import {
  BUNDLED_FONTS,
  createFontProvider,
  loadBundledFontBytes,
  maxDigitWidth,
  resolveBundledFamilyFace,
  resolveDigitAdvance,
  resolveLastResortFace,
  resolveMetricCompatFace,
  resolveMetricCompatFamily,
  resolveScriptFallbackFace,
} from './index';

const SFNT_TRUETYPE = 0x00010000;
const SFNT_CFF = 0x4f54544f;

describe('resolution', () => {
  test('maps the MS core fonts to their metric-compatible substitutes', () => {
    expect(resolveMetricCompatFamily('Calibri')).toBe('Carlito');
    expect(resolveMetricCompatFamily('  CAMBRIA ')).toBe('Caladea');
    expect(resolveMetricCompatFamily('Helvetica')).toBe('Liberation Sans');
    expect(resolveMetricCompatFamily('Times')).toBe('Liberation Serif');
    expect(resolveMetricCompatFamily('Nonesuch')).toBeUndefined();
  });

  test('falls back to a family regular when the exact style is not vendored', () => {
    expect(resolveMetricCompatFace('Calibri', true, true)?.file).toBe('Carlito-BoldItalic.ttf');
    expect(resolveMetricCompatFace('SimSun', true, false)?.file).toBe('NotoSerifSC-Regular.otf');
  });

  test('last resort always returns a face, serif-aware', () => {
    expect(resolveLastResortFace('Totally Unknown', false, false).family).toBe('Carlito');
    expect(resolveLastResortFace('Garamond', false, false).family).toBe('Liberation Serif');
    expect(resolveLastResortFace('Unknown', true, true).file).toBe('Carlito-BoldItalic.ttf');
    expect(resolveLastResortFace('Calibri Light', false, false).file).toBe('Carlito-Regular.ttf');
    expect(resolveLastResortFace(' CALIBRI LIGHT ', true, true).file).toBe('Carlito-Italic.ttf');
    expect(resolveMetricCompatFamily('Calibri Light')).toBeUndefined();
  });

  test('Word substitutes Arial for an unknown sans family, PowerPoint Calibri', () => {
    expect(resolveLastResortFace('Lato', false, false, 'word').file).toBe('LiberationSans-Regular.ttf');
    expect(resolveLastResortFace('Lato', true, false, 'powerpoint').file).toBe('Carlito-Bold.ttf');
    expect(resolveLastResortFace('Garamond', false, false, 'word').family).toBe('Liberation Serif');
    expect(resolveLastResortFace('Consolas', false, false, 'word').file).toBe('LiberationMono-Regular.ttf');
  });

  test('a typewriter or old-style name lands on a face of its own kind', () => {
    expect(resolveLastResortFace('Consolas', false, false).file).toBe('LiberationMono-Regular.ttf');
    expect(resolveLastResortFace('Lucida Console', false, false).file).toBe(
      'LiberationMono-Regular.ttf'
    );
    expect(resolveLastResortFace('Book Antiqua', false, false).file).toBe(
      'LiberationSerif-Regular.ttf'
    );
    expect(resolveLastResortFace('Bookman Old Style', true, false).file).toBe(
      'LiberationSerif-Bold.ttf'
    );
    expect(resolveLastResortFace('Calisto MT', false, false).file).toBe(
      'LiberationSerif-Regular.ttf'
    );
    expect(resolveLastResortFace('Monotype Corsiva', false, false).file).toBe(
      'Carlito-Regular.ttf'
    );
  });

  test('a family with no clone takes the bundled face closest in width', () => {
    expect(resolveLastResortFace('Verdana', false, false).file).toBe('Montserrat-Regular.ttf');
    expect(resolveLastResortFace('Verdana', true, false).file).toBe('Montserrat-Bold.ttf');
    expect(resolveLastResortFace('Trebuchet MS', false, false).file).toBe('Carlito-Regular.ttf');
    expect(resolveLastResortFace('Gill Sans MT', false, false).file).toBe('Carlito-Regular.ttf');
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
    expect(resolveLastResortFace('Calibri Light', true, false).file).toBe('Carlito-Regular.ttf');
    expect(resolveLastResortFace('Inter Light', true, false).file).toBe('Inter-Regular.ttf');
    expect(resolveLastResortFace('Archivo Black', true, false).file).toBe('Carlito-Bold.ttf');
    expect(resolveLastResortFace('Lato Light', true, false).file).toBe('Carlito-Bold.ttf');
    expect(resolveLastResortFace('Blackadder ITC', true, false).file).toBe('Carlito-Bold.ttf');
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
    expect(resolveScriptFallbackFace('hebrew', true, false)?.file).toBe('NotoSansHebrew-Bold.ttf');
  });

  test('known heavy family names select the bold face and preserve italics', () => {
    for (const family of ['Arial Black', ' ARIAL-BLACK ', 'Arial Heavy', 'Arial ExtraBold']) {
      expect(resolveLastResortFace(family, false, false).file).toBe('LiberationSans-Bold.ttf');
      expect(resolveLastResortFace(family, false, true).file).toBe('LiberationSans-BoldItalic.ttf');
    }
    expect(resolveLastResortFace('Calibri Heavy', false, false).file).toBe('Carlito-Bold.ttf');
    expect(resolveLastResortFace('Calibri Heavy', false, true).file).toBe('Carlito-BoldItalic.ttf');
  });

  test('unknown heavy family names keep the normal fallback weight', () => {
    expect(resolveLastResortFace('Archivo Black', false, false).file).toBe('Carlito-Regular.ttf');
    expect(resolveLastResortFace('Archivo Black', false, true).file).toBe('Carlito-Italic.ttf');
    expect(resolveLastResortFace('Archivo Black', true, false).file).toBe('Carlito-Bold.ttf');
  });
});

describe('bundled family names', () => {
  test('resolve to their own faces, not only as Word families', () => {
    expect(resolveMetricCompatFace('Gelasio', false, false)).toBeUndefined();
    expect(resolveBundledFamilyFace('Gelasio', true, true)?.file).toBe('Gelasio-BoldItalic.ttf');
    expect(resolveBundledFamilyFace(' source sans 3 ', false, true)?.file).toBe(
      'SourceSans3-Italic.ttf'
    );
    expect(resolveBundledFamilyFace('Comic Relief', true, false)?.file).toBe('ComicRelief-Bold.ttf');
    expect(resolveBundledFamilyFace('Noto Serif SC', false, false)?.file).toBe(
      'NotoSerifSC-Regular.otf'
    );
  });

  test('never stand in another style, so the browser can synthesize it', () => {
    expect(resolveBundledFamilyFace('Oswald', false, true)).toBeUndefined();
    expect(resolveBundledFamilyFace('Noto Serif SC', true, false)).toBeUndefined();
    expect(resolveBundledFamilyFace('SimSun', true, false)).toBeUndefined();
    expect(resolveMetricCompatFace('SimSun', true, false)?.file).toBe('NotoSerifSC-Regular.otf');
  });

  test('cover the serif Noto families the package does not vendor with the sans face', () => {
    expect(resolveBundledFamilyFace('Noto Serif TC', false, false)?.family).toBe('Noto Sans TC');
    expect(resolveBundledFamilyFace('Noto Serif JP', false, false)?.family).toBe('Noto Sans JP');
    expect(resolveBundledFamilyFace('Noto Serif KR', false, false)?.family).toBe('Noto Sans KR');
  });

  test('fall back to Word-family resolution and miss unknown names', () => {
    expect(resolveBundledFamilyFace('Calibri', true, true)?.file).toBe('Carlito-BoldItalic.ttf');
    expect(resolveBundledFamilyFace('Tinos', false, false)?.family).toBe('Liberation Serif');
    expect(resolveBundledFamilyFace('Nonesuch', false, false)).toBeUndefined();
  });

  test('are served by the provider', async () => {
    const bytes = await createFontProvider().resolveFamily('Gelasio', false, false)!();
    expect(bytes.byteLength).toBe(107588);
    expect(createFontProvider().resolve('Gelasio', false, false)).toBeUndefined();
  });
});

describe('loading', () => {
  test('createFontProvider serves real sfnt bytes', async () => {
    const provider = createFontProvider();
    const bytes = await provider.resolve('Calibri', false, false)!();
    expect(new DataView(bytes).getUint32(0)).toBe(SFNT_TRUETYPE);
    expect(bytes.byteLength).toBe(628032);
  });

  test('the CJK add-on resolves through the optional dynamic import', async () => {
    const provider = createFontProvider();
    const bytes = await provider.resolveScriptFallback('cjk-sc', false, false)!();
    expect(new DataView(bytes).getUint32(0)).toBe(SFNT_CFF);
    const { script: _script, ...withoutScript } = resolveScriptFallbackFace('cjk-sc', false, false)!;
    expect(new DataView(await loadBundledFontBytes(withoutScript)).getUint32(0)).toBe(SFNT_CFF);
  });

  test('caches per face, handing out one buffer identity', async () => {
    const face = resolveMetricCompatFace('Cambria', false, false)!;
    const [a, b] = await Promise.all([loadBundledFontBytes(face), loadBundledFontBytes(face)]);
    expect(a).toBe(b);
  });

  test('pins a relative provider base to the route where it is created', async () => {
    const originalLocation = Object.getOwnPropertyDescriptor(globalThis, 'location');
    const realFetch = globalThis.fetch;
    const face = resolveMetricCompatFace('Calibri', false, true)!;
    const expected = await Bun.file(new URL(`../assets/${face.file}`, import.meta.url)).arrayBuffer();
    let requested = '';
    Object.defineProperty(globalThis, 'location', {
      configurable: true,
      value: new URL('https://example.test/docs/team/report/'),
    });
    try {
      globalThis.fetch = (async (input: RequestInfo | URL) => {
        requested = String(input);
        return new Response(expected);
      }) as unknown as typeof fetch;
      const provider = createFontProvider({ baseUrl: 'fonts/' });
      Object.defineProperty(globalThis, 'location', {
        configurable: true,
        value: new URL('https://example.test/'),
      });

      const bytes = await provider.resolve('Calibri', false, true)!();

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
    const face = resolveMetricCompatFace('Calibri', true, true)!;
    const intact = await Bun.file(new URL(`../assets/${face.file}`, import.meta.url)).arrayBuffer();
    let attempts = 0;
    globalThis.fetch = (async () => {
      attempts++;
      return new Response(
        attempts === 1 ? intact.slice(0, Math.floor(intact.byteLength / 2)) : intact
      );
    }) as unknown as typeof fetch;
    try {
      const options = { baseUrl: 'https://length-check.example/assets/' };
      await expect(loadBundledFontBytes(face, options)).rejects.toThrow(
        `expected ${face.byteLength}`
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
      expect(bytes.byteLength).toBe(319508);
    } finally {
      globalThis.fetch = realFetch;
    }
  });

  test('loads every manifest face from the package without a base URL', async () => {
    for (const face of BUNDLED_FONTS) {
      expect((await loadBundledFontBytes(face)).byteLength).toBe(face.byteLength);
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

  test('gives every shipped asset a literal URL a bundler can emit', async () => {
    const source = await Bun.file(new URL('./index.ts', import.meta.url)).text();
    const rows = [
      ...source.matchAll(/'([^']+)': \(\) =>\s+new URL\('\.\.\/assets\/([^']+)', import\.meta\.url\)/g),
    ];
    const assets = readdirSync(new URL('../assets/', import.meta.url))
      .filter((file) => /\.(ttf|otf)$/.test(file))
      .sort();
    expect(rows.map(([, file]) => file).sort()).toEqual(assets);
    expect(rows.filter(([, file, target]) => file !== target).map(([row]) => row)).toEqual([]);
    expect(
      BUNDLED_FONTS.filter((face) => !face.script?.startsWith('cjk-'))
        .map((face) => face.file)
        .sort()
    ).toEqual(assets);
  });

  test('records the exact byte length of every vendored face', () => {
    for (const face of BUNDLED_FONTS) {
      const directory = face.script?.startsWith('cjk-') ? '../../fonts-cjk/assets/' : '../assets/';
      expect(Bun.file(new URL(`${directory}${face.file}`, import.meta.url)).size).toBe(
        face.byteLength
      );
    }
  });
});

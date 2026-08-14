import { describe, expect, test } from 'bun:test';
import { existsSync } from 'node:fs';

import manifest from '../package.json' with { type: 'json' };
import {
  BUNDLED_FONTS,
  createFontProvider,
  loadBundledFontBytes,
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
    expect(resolveLastResortFace('Totally Unknown', false, false).family).toBe('Liberation Sans');
    expect(resolveLastResortFace('Garamond', false, false).family).toBe('Liberation Serif');
    expect(resolveLastResortFace('Unknown', true, true).file).toBe('LiberationSans-BoldItalic.ttf');
  });

  test('script fallbacks prefer the sans face of the bucket', () => {
    expect(resolveScriptFallbackFace('cjk-sc', false, false)?.family).toBe('Noto Sans SC');
    expect(resolveScriptFallbackFace('arabic', false, false)?.family).toBe('Noto Sans Arabic');
    expect(resolveScriptFallbackFace('hebrew', true, false)?.file).toBe('NotoSansHebrew-Bold.ttf');
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

  test('records the exact byte length of every vendored face', () => {
    for (const face of BUNDLED_FONTS) {
      const directory = face.script?.startsWith('cjk-') ? '../../fonts-cjk/assets/' : '../assets/';
      expect(Bun.file(new URL(`${directory}${face.file}`, import.meta.url)).size).toBe(
        face.byteLength
      );
    }
  });
});

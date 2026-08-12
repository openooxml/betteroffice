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
const SFNT_CFF = 0x4f54544f; // 'OTTO'

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
    // CJK ships Regular only; bold resolves to it rather than to nothing.
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

  /**
   * Node's `fetch` does not implement `file:` — only Bun's does — so the
   * loader must read `file:` assets off disk. Stubbing `fetch` to throw is the
   * only way to prove that from Bun, where a passing `fetch` would mask it.
   */
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

/**
 * Guards the published shape. Pointing `main`/`types` at `src/*.ts` makes the
 * package unimportable from plain Node, which refuses to type-strip anything
 * under `node_modules` (ERR_UNSUPPORTED_NODE_MODULES_TYPE_STRIPPING).
 */
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
});

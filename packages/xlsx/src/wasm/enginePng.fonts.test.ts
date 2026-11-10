import { afterEach, beforeEach, describe, expect, mock, test } from 'bun:test';
import type { XlsxPngEngine } from './enginePng';
import type { ResolvedPngTarget } from './pngResolve';

const painted: unknown[] = [];
let loaded = 0;
const bundled = [{ family: 'Bundled', bytes: new Uint8Array([1]) }];

beforeEach(() => {
  painted.length = 0;
  loaded = 0;
  mock.module('../render/png', () => ({
    canvasPngAvailable: () => true,
    canvasUnavailable: () => new Error('no canvas'),
    sheetFontRequests: () => [],
    displayListToPng: (_list: unknown, options: { fonts?: unknown; fallbacks?: unknown }) => {
      painted.push(options);
      return Promise.resolve(new Uint8Array());
    },
  }));
  mock.module('./pngResolve', () => ({
    pngDisplayList: () => ({ list: { commands: [] }, width: 1, height: 1, dpr: 1 }),
  }));
  mock.module('./enginePngFonts', () => ({
    defaultSheetFaces: () => {
      loaded += 1;
      return Promise.resolve({ faces: bundled, fallbacks: () => [] });
    },
  }));
});

afterEach(() => mock.restore());

async function render(options: Parameters<typeof import('./enginePng').renderPng>[2]) {
  const { renderPng } = await import('./enginePng');
  await renderPng({} as XlsxPngEngine, {} as ResolvedPngTarget, options);
  return painted[painted.length - 1] as { fonts: unknown; fallbacks?: unknown };
}

const nativeEngine = (globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] === 'native';

describe.skipIf(nativeEngine)('wasm PNG fonts option', () => {
  test('an absent list paints with the bundled faces', async () => {
    const options = await render({});
    expect(loaded).toBe(1);
    expect(options.fonts).toBe(bundled);
  });

  test('an explicit empty list paints with no bundled faces', async () => {
    const options = await render({ fonts: [] });
    expect(loaded).toBe(0);
    expect(options.fonts).toEqual([]);
    expect(options.fallbacks).toBeUndefined();
  });
});

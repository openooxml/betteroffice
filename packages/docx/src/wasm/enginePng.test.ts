import { afterEach, describe, expect, test } from 'bun:test';
import type { DisplayList, DisplayPage } from '../layout/render/displayList';
import { engineKind, renderPng, type DocxPngEngine } from './enginePng';

const page = (primitives: unknown[], size = { width: 200, height: 100 }) => ({ pageIndex: 0, ...size, primitives }) as unknown as DisplayPage;
const picture = { kind: 'image', x: 0, y: 0, width: 10, height: 10, relId: 'media:logo' };
const glyphs = { kind: 'glyphRun', x: 0, y: 0, fontId: 1, size: 12, color: '#000000', text: 'A', glyphs: [{ id: 5, x: 0, y: 0, advance: 6 }] };
const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');

const engine = (drawn: DisplayPage, overrides: Partial<DocxPngEngine> = {}): DocxPngEngine => ({
  displayList: { pages: [drawn] } as DisplayList, fontChains: {}, page: drawn, session: null,
  outlineGlyph: () => JSON.stringify({ commands: [] }),
  media: () => ({ bytes: svg, mimeType: 'image/svg+xml' }),
  ...overrides,
});

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stub(canvas = true) {
  const calls = { sizes: [] as number[][], encoded: 0, backgrounds: [] as unknown[] };
  class Canvas {
    constructor(public width: number, public height: number) { calls.sizes.push([width, height]); }
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'fillRect') return () => { calls.backgrounds.push(target.fillStyle); };
          if (property === 'measureText') return (text: string) => ({ width: text.length });
          return property in target ? target[property as string] : () => undefined;
        },
        set(target, property, value) { target[property as string] = value; return true; },
      });
    }
    convertToBlob() { calls.encoded++; return Promise.resolve(new Blob(['png'])); }
  }
  class Path { moveTo() {} lineTo() {} quadraticCurveTo() {} bezierCurveTo() {} closePath() {} }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, Path2D: scope.Path2D, Image: scope.Image, document: scope.document };
  Object.assign(scope, { OffscreenCanvas: canvas ? Canvas : undefined, Path2D: Path, Image: undefined, document: undefined });
  restore = () => { Object.assign(scope, saved); };
  return calls;
}

describe.skipIf(engineKind === 'native')('canvas renderPng', () => {
  test('paints white behind a page without a background, at the scaled size', async () => {
    const calls = stub();
    expect(new TextDecoder().decode(await renderPng(engine(page([])), { scale: 1.5 }))).toBe('png');
    expect(calls.sizes).toEqual([[300, 150]]);
    expect(calls.backgrounds[0]).toBe('#ffffff');
  });

  test('an SVG picture where only a document can decode it rejects instead of leaving it out', async () => {
    const calls = stub();
    await expect(renderPng(engine(page([picture])), {})).rejects.toMatchObject({
      name: 'PngExportError', code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/docx-native'),
    });
    expect(calls.encoded).toBe(0);
  });

  test('a glyph whose outline fails rejects instead of painting browser text', async () => {
    const calls = stub();
    const failing = engine(page([glyphs]), { outlineGlyph: () => { throw new Error('no outline'); } });
    await expect(renderPng(failing, {})).rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
    expect(calls.encoded).toBe(0);
  });

  test('refuses an oversize page and a runtime without a canvas', async () => {
    stub();
    await expect(renderPng(engine(page([], { width: 20_000, height: 100 })), {})).rejects.toMatchObject({ code: 'canvas-too-large' });
    restore?.();
    stub(false);
    await expect(renderPng(engine(page([])), {})).rejects.toMatchObject({ code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/docx-native') });
  });

  test('stops without encoding when aborted', async () => {
    const calls = stub();
    const controller = new AbortController();
    controller.abort();
    await expect(renderPng(engine(page([])), { signal: controller.signal })).rejects.toBeDefined();
    expect(calls.encoded).toBe(0);
  });
});

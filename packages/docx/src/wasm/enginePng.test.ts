import { afterEach, describe, expect, test } from 'bun:test';
import type { PngDiagnostic } from '../../../../shared/office-session/pngExportError';
import type { DisplayList, DisplayPage } from '../layout/render/displayList';
import { engineKind, renderPng, type DocxPngEngine } from './enginePng';

const page = (primitives: unknown[], size = { width: 200, height: 100 }) => ({ pageIndex: 0, ...size, primitives }) as unknown as DisplayPage;
const picture = { kind: 'image', x: 0, y: 0, w: 10, h: 10, relId: 'media:logo' };
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

function stub(canvas = true, context = true) {
  const calls = { sizes: [] as number[][], encoded: 0, backgrounds: [] as unknown[], fills: [] as unknown[], drawn: 0 };
  class Canvas {
    constructor(public width: number, public height: number) { calls.sizes.push([width, height]); }
    getContext() {
      if (!context) return null;
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'fillRect') return () => { calls.backgrounds.push(target.fillStyle); };
          if (property === 'fill') return () => { calls.fills.push(target.fillStyle); };
          if (property === 'drawImage') return () => { calls.drawn++; };
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
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, Path2D: scope.Path2D, Image: scope.Image, document: scope.document, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: canvas ? Canvas : undefined, Path2D: Path, Image: undefined, document: undefined,
    createImageBitmap: (blob: Blob) => blob.type === 'image/png' ? Promise.resolve({ width: 4, height: 4, close() {} }) : Promise.reject(new Error('undecodable')),
  });
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

  test('an SVG picture only a document can decode paints a placeholder and is reported', async () => {
    const calls = stub();
    const diagnostics: PngDiagnostic[] = [];
    expect(new TextDecoder().decode(await renderPng(engine(page([picture])), { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) }))).toBe('png');
    expect(calls.encoded).toBe(1);
    expect(calls.fills).toEqual(['rgb(242, 242, 242)']);
    expect(diagnostics).toEqual([{ code: 'picture-unavailable', pictures: [{ id: 'media:logo', reason: 'unsupported' }] }]);
  });

  test('a picture whose part cannot be read paints a placeholder, reported as missing; a data URL is named by its header', async () => {
    const calls = stub();
    const diagnostics: PngDiagnostic[] = [];
    const data = { ...picture, relId: 'data:image/svg+xml;base64,PHN2Zy8+' };
    await renderPng(engine(page([picture, data]), { media: () => null }), { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    expect(calls.fills).toEqual(['rgb(242, 242, 242)', 'rgb(242, 242, 242)']);
    expect(diagnostics).toEqual([{ code: 'picture-unavailable', pictures: [
      { id: 'media:logo', reason: 'missing' }, { id: 'data:image/svg+xml;base64', reason: 'unsupported' },
    ] }]);
  });

  test('a page whose pictures all draw reports nothing', async () => {
    const calls = stub();
    const diagnostics: PngDiagnostic[] = [];
    const png = { bytes: new Uint8Array([137, 80, 78, 71]), mimeType: 'image/png' };
    await renderPng(engine(page([picture]), { media: () => png }), { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    expect(calls.drawn).toBe(1);
    expect(calls.fills).toEqual([]);
    expect(diagnostics).toEqual([]);
  });

  test('a raster the decoder refuses paints a placeholder, reported as decode-failed', async () => {
    const calls = stub();
    const diagnostics: PngDiagnostic[] = [];
    await renderPng(engine(page([picture]), { media: () => ({ bytes: new Uint8Array([1, 2, 3]), mimeType: 'image/jpeg' }) }), { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    expect(calls.drawn).toBe(0);
    expect(calls.fills).toEqual(['rgb(242, 242, 242)']);
    expect(diagnostics).toEqual([{ code: 'picture-unavailable', pictures: [{ id: 'media:logo', reason: 'decode-failed' }] }]);
  });

  test('the stand-in the parser shows for an unconvertible metafile paints the placeholder, reported as metafile', async () => {
    const calls = stub();
    const diagnostics: PngDiagnostic[] = [];
    const standIn = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="96" height="72.5" viewBox="0 0 96 72.5" preserveAspectRatio="none"><rect x="0.5" y="0.5" width="95" height="71.5" fill="#f1f3f4" stroke="#c4c7c5"/><g transform="translate(48 36.25) scale(1.5104)" fill="none" stroke="#9aa0a6" stroke-width="1.5" stroke-linejoin="round"><rect x="-12" y="-9" width="24" height="18" rx="2"/><circle cx="-5" cy="-3.5" r="2"/><path d="M-10 7L-3 0L2 5L5 2L10 7"/></g></svg>');
    await renderPng(engine(page([picture]), { media: () => ({ bytes: standIn, mimeType: 'image/svg+xml' }) }), { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    expect(calls.fills).toEqual(['rgb(242, 242, 242)']);
    expect(diagnostics).toEqual([{ code: 'picture-unavailable', pictures: [{ id: 'media:logo', reason: 'metafile' }] }]);
  });

  test('a picture fill without a reference paints a placeholder instead of its solid fill', async () => {
    const calls = stub();
    const diagnostics: PngDiagnostic[] = [];
    const shape = {
      kind: 'shape', x: 0, y: 0, w: 10, h: 10, fill: '#ff0000',
      geometryPath: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 10, y: 0 }, { type: 'line', x: 10, y: 10 }, { type: 'close' }],
      fillPaint: { kind: 'picture' },
    };
    await renderPng(engine(page([shape])), { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    expect(calls.fills).toEqual(['rgb(242, 242, 242)']);
    expect(diagnostics).toEqual([{ code: 'picture-unavailable', pictures: [{ id: '', reason: 'missing' }] }]);
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

  test('a canvas without a 2d context, as in Deno, rejects as unavailable', async () => {
    const calls = stub(true, false);
    await expect(renderPng(engine(page([])), {})).rejects.toMatchObject({ name: 'PngExportError', code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/docx-native') });
    expect(calls.encoded).toBe(0);
  });

  test('stops without encoding when aborted', async () => {
    const calls = stub();
    const controller = new AbortController();
    controller.abort();
    await expect(renderPng(engine(page([])), { signal: controller.signal })).rejects.toBeDefined();
    expect(calls.encoded).toBe(0);
  });
});

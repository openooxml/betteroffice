import { afterEach, describe, expect, test } from 'bun:test';
import type { SlideDisplayList, SlidePrimitive } from '../types';
import { BROWSER_FONT_ID, pngBackground, slideFace, slideFonts, slideListToPng } from './png';

const card: SlidePrimitive = {
  kind: 'shape', objectId: 1, shapeId: 'shape:1', name: 'Card', x: 20, y: 20, w: 280, h: 140, geometry: 'rect',
  path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 0 }, { type: 'line', x: 1, y: 1 }, { type: 'line', x: 0, y: 1 }, { type: 'close' }],
  fill: { kind: 'solid', color: '#3b82f6' },
} as SlidePrimitive;

const run = (family: string, bold = false, fontId = 1) => ({
  text: 'Hello', start: 0, end: 5, x: 40, width: 60, fontId, fontFamily: family, fontSizePx: 20,
  bold, italic: false, underline: false, color: '#000000', glyphs: [],
});

const textBox = (family: string, bold = false, fontId = 1): SlidePrimitive => ({
  kind: 'textBox', objectId: 2, shapeId: 'shape:2', storyId: 'story:2', x: 40, y: 50, w: 240, h: 80, anchor: 'top', paragraphs: [],
  lines: [{ x: 40, y: 50, width: 60, height: 24, baseline: 68, start: 0, end: 5, caretStops: [], runs: [run(family, bold, fontId)] }],
} as unknown as SlidePrimitive);

const list = (primitives: SlidePrimitive[] = [card]): SlideDisplayList => ({
  contractVersion: 1, width: 320, height: 180, background: { kind: 'solid', color: '#ffffff' }, primitives,
} as SlideDisplayList);

interface Calls { sizes: [number, number][]; fonts: string[]; faces: string[]; held: unknown[]; encoded: number; ops: string[] }

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stub(options: { faceFails?: boolean; canvas?: boolean } = {}): Calls {
  const calls: Calls = { sizes: [], fonts: [], faces: [], held: [], encoded: 0, ops: [] };
  class Canvas {
    constructor(width: number, height: number) { calls.sizes.push([width, height]); }
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'createLinearGradient' || property === 'createRadialGradient') return () => ({ addColorStop: () => undefined });
          if (property === 'measureText') return (text: string) => ({ width: text.length });
          if (property === 'fillRect') return () => { calls.ops.push(`fillRect:${String(target.globalCompositeOperation)}:${String(target.fillStyle)}`); };
          if (property in target) return target[property as string];
          return () => undefined;
        },
        set(target, property, value) {
          if (property === 'font') calls.fonts.push(String(value));
          target[property as string] = value;
          return true;
        },
      });
    }
    convertToBlob(request: { type: string }) { calls.encoded++; return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  class Face {
    family: string;
    constructor(family: string) { this.family = family; calls.faces.push(family); }
    load() { return options.faceFails ? Promise.reject(new Error('bad font')) : Promise.resolve(this); }
  }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, FontFace: scope.FontFace, fonts: scope.fonts };
  scope.OffscreenCanvas = options.canvas === false ? undefined : Canvas;
  scope.FontFace = Face;
  scope.fonts = {
    add(face: unknown) { calls.held.push(face); },
    delete(face: unknown) { calls.held.splice(calls.held.indexOf(face), 1); return true; },
    load(font: string) { return Promise.resolve(calls.held.filter((face) => font.includes(`"${(face as Face).family}"`))); },
  };
  restore = () => { Object.assign(scope, saved); };
  return calls;
}

describe('slideListToPng', () => {
  test('sizes the canvas as the rasterizer does, scale and fractional extents included, and fills white behind the slide', async () => {
    const calls = stub();
    expect(new TextDecoder().decode(await slideListToPng(list()))).toBe('png');
    await slideListToPng(list(), { scale: 2 });
    await slideListToPng({ ...list(), width: 1122.1667, height: 793.2 });
    expect(calls.sizes).toEqual([[320, 180], [640, 360], [1123, 794]]);
    expect(calls.ops.filter((op) => op.startsWith('fillRect:destination-over'))).toEqual(Array(3).fill('fillRect:destination-over:#ffffff'));
  });

  test('backgrounds read like the native renderer: slide, transparent or a #rrggbb color', async () => {
    const calls = stub();
    await slideListToPng(list(), { background: 'transparent' });
    await slideListToPng(list(), { background: '#102030' });
    expect(calls.ops.filter((op) => op.startsWith('fillRect:destination-over'))).toEqual(['fillRect:destination-over:#102030']);
    expect(() => pngBackground('red')).toThrow('background must be "slide", "transparent", or a #rrggbb color');
  });

  test('refuses a bad scale, an oversize slide and a runtime without a canvas', async () => {
    const calls = stub();
    await expect(slideListToPng(list(), { scale: 0 })).rejects.toThrow('finite and positive');
    await expect(slideListToPng(list(), { scale: 100 })).rejects.toMatchObject({ name: 'PngExportError', code: 'canvas-too-large' });
    expect(calls.sizes).toEqual([]);
    restore?.();
    stub({ canvas: false });
    await expect(slideListToPng(list())).rejects.toMatchObject({ code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/pptx-native') });
  });

  test('paints text with a supplied face under a private family and drops it afterwards', async () => {
    const calls = stub();
    await slideListToPng(list([textBox('Brand Sans', true), textBox('Other', false, BROWSER_FONT_ID)]), { fonts: [{ family: 'brand sans', bytes: new Uint8Array([1]) }] });
    expect(calls.faces).toEqual([expect.stringMatching(/^bo-pptx-png-\d+-0$/)]);
    expect(calls.fonts.some((font) => /^bold 20px bo-pptx-png-\d+-0$/.test(font))).toBe(true);
    expect(calls.fonts.some((font) => font.endsWith('20px Other'))).toBe(true);
    expect(calls.held).toEqual([]);
  });

  test('a supplied face that cannot load rejects instead of painting a substitute', async () => {
    const calls = stub({ faceFails: true });
    await expect(slideListToPng(list([textBox('Brand Sans')]), { fonts: [{ family: 'Brand Sans', bytes: new Uint8Array([1]) }] }))
      .rejects.toMatchObject({ code: 'font-unavailable', message: expect.stringContaining('Brand Sans') });
    expect(calls.encoded).toBe(0);
  });

  test('text laid out with an engine face the export has no copy of rejects instead of painting by family name', async () => {
    const calls = stub();
    await expect(slideListToPng(list([textBox('Brand Sans', false, 7)]), { fonts: [{ id: 3, family: 'Brand Sans', bytes: new Uint8Array([1]) }] }))
      .rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable', message: expect.stringContaining('Brand Sans') });
    expect([calls.faces, calls.encoded]).toEqual([[], 0]);
  });

  test('stops without encoding when aborted', async () => {
    const calls = stub();
    const controller = new AbortController();
    controller.abort();
    await expect(slideListToPng(list(), { signal: controller.signal })).rejects.toBeDefined();
    expect(calls.encoded).toBe(0);
  });
});

describe('slide fonts', () => {
  test('collect text inside tables and charts, and choose a face by family dropping bold before italic', () => {
    const table = { kind: 'table', objectId: 3, shapeId: 'shape:3', name: 't', label: 't', x: 0, y: 0, w: 1, h: 1, primitives: [textBox('Cell Sans')] } as unknown as SlidePrimitive;
    expect(slideFonts(list([table, textBox('Brand Sans', true)]))).toEqual([
      { family: 'Cell Sans', bold: false, italic: false, id: 1 }, { family: 'Brand Sans', bold: true, italic: false, id: 1 },
    ]);
    const regular = { family: 'Brand Sans', bytes: new Uint8Array([1]) };
    const italic = { family: 'Brand Sans', italic: true, bytes: new Uint8Array([2]) };
    expect(slideFace([regular, italic], { family: 'Brand Sans', bold: true, italic: true })).toBe(italic);
    expect(slideFace([regular], { family: 'Other', bold: false, italic: false })).toBeUndefined();
    const measured = { id: 3, family: 'Liberation Sans', bytes: new Uint8Array([3]) };
    expect(slideFace([regular, measured], { family: 'Calibri', bold: false, italic: false, id: 3 })).toBe(measured);
    expect(slideFace([measured], { family: 'Liberation Sans', bold: false, italic: false, id: 4 })).toBeUndefined();
  });
});

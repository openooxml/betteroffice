import { afterEach, describe, expect, test } from 'bun:test';
import type { PageDisplayList, TextBoxPrimitive, VsdxFontFace } from '../types';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { paintPage } from './canvas';
import { pageToPng } from './png';

const run = { text: 'Hello', family: 'Brand Sans', sizeIn: 12, bold: false, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#000000' };
const textBox: TextBoxPrimitive = { id: 't', zOrder: 1, kind: 'textBox', x: 0, y: 0, width: 100, height: 20, paragraphs: [{ runs: [run] }], lines: [{ x: 0, y: 0, width: 100, height: 20, start: 0, end: 5, caretStops: [] }] };
const list = (primitives: PageDisplayList['primitives'] = [textBox]) => ({
  contractVersion: 7, width: 100.5, height: 50, primitives, paintTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
}) as unknown as PageDisplayList;
const face: VsdxFontFace = { family: 'Brand Sans', bytes: new Uint8Array([1, 2, 3]) };

interface Calls { sizes: [number, number][]; ctx: string[]; encoded: number; fonts: string[]; }

function stubBrowser(options: { failLoad?: boolean; held?: { family: string; weight: string; style: string; status: string; load(): Promise<unknown> }[] } = {}) {
  const calls: Calls = { sizes: [], ctx: [], encoded: 0, fonts: [] };
  const held = [...(options.held ?? [])];
  const loads: string[] = [];
  const set = {
    [Symbol.iterator]: () => held[Symbol.iterator](),
    add(value: never) { calls.ctx.push('fonts.add'); held.push(value); },
    delete(value: never) { calls.ctx.push('fonts.delete'); held.splice(held.indexOf(value), 1); return true; },
    load(font: string) { loads.push(font); return Promise.resolve(held.filter((value) => font.endsWith(` ${value.family}`))); },
  };
  class StubFace {
    status = 'unloaded'; family: string; weight: string; style: string;
    constructor(family: string, _source: ArrayBuffer, descriptors: { weight: string; style: string }) {
      this.family = family; this.weight = descriptors.weight; this.style = descriptors.style;
      calls.ctx.push('FontFace');
    }
    load() {
      if (options.failLoad) return Promise.reject(new Error('bad font'));
      this.status = 'loaded';
      return Promise.resolve(this);
    }
  }
  class StubCanvas {
    constructor(width: number, height: number) { calls.sizes.push([width, height]); }
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'measureText') return () => ({ width: 1 });
          if (property === 'createLinearGradient' || property === 'createRadialGradient') return () => ({ addColorStop: () => undefined });
          if (property in target) return target[property as string];
          return () => { calls.ctx.push(String(property)); };
        },
        set(target, property, value) { if (property === 'font') calls.fonts.push(String(value)); target[property as string] = value; return true; },
      });
    }
    convertToBlob(request: { type: string }) { calls.encoded += 1; return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, FontFace: scope.FontFace, fonts: scope.fonts, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, { OffscreenCanvas: StubCanvas, FontFace: StubFace, fonts: set });
  undo = () => { Object.assign(scope, saved); };
  return { calls, loads, held };
}

let undo: (() => void) | undefined;
afterEach(() => { undo?.(); undo = undefined; });

describe('pageToPng', () => {
  test('rounds the page size up like the raster backend and fills white behind the page', async () => {
    const { calls } = stubBrowser();
    const png = await pageToPng(list([]), { scale: 2 });
    expect(new TextDecoder().decode(png)).toBe('png');
    expect(calls.sizes).toEqual([[201, 100]]);
    expect(calls.ctx[calls.ctx.length - 1]).toBe('restore');
    expect(calls.ctx).toContain('fillRect');
  });

  test('installs a face the page uses under its own name before painting and removes it afterwards', async () => {
    const { calls, loads, held } = stubBrowser();
    await pageToPng(list(), { fonts: [face, { family: 'Unused', bytes: new Uint8Array([1]) }] });
    expect(calls.ctx.filter(call => call === 'FontFace').length).toBe(1);
    expect(calls.ctx.indexOf('fonts.add')).toBeLessThan(calls.ctx.indexOf('fillText'));
    expect(loads).toHaveLength(1);
    expect(loads[0]).toMatch(/^16px bo-vsdx-png-\d+-0$/);
    expect(calls.fonts.length).toBeGreaterThan(0);
    expect(calls.fonts.every((font) => /bo-vsdx-png-\d+-0$/.test(font))).toBe(true);
    expect(held.length).toBe(0);
  });

  test('never paints with a face the editor or another export installed under the same family', async () => {
    const existing = { family: 'Brand Sans', weight: '400', style: 'normal', status: 'loaded', load() { return Promise.resolve(this); } };
    const { calls, held } = stubBrowser({ held: [existing] });
    let finishImage!: (value: unknown) => void;
    const imageReady = new Promise((resolve) => { finishImage = resolve; });
    (globalThis as { createImageBitmap?: unknown }).createImageBitmap = () => imageReady.then(() => ({ close() {} }));
    const image = { id: 'i', zOrder: 0, kind: 'image', assetId: 'a', x: 0, y: 0, width: 1, height: 1 };
    const slow = pageToPng(list([image as never, textBox]), { fonts: [face], media: () => new Uint8Array([1]) });
    await pageToPng(list(), { fonts: [{ family: 'Brand Sans', bytes: new Uint8Array([9, 9]) }] });
    finishImage(undefined);
    await slow;
    const families = new Set(calls.fonts.map((font) => font.replace(/^.*px /, '')));
    expect(families.size).toBe(2);
    for (const family of families) expect(family).toMatch(/^bo-vsdx-png-\d+-0$/);
    expect(held).toEqual([existing]);
  });

  test('a quoted multiword family resolves to its registered face', async () => {
    const { calls } = stubBrowser();
    const box = { ...textBox, paragraphs: [{ runs: [{ ...run, family: '"Brand Sans"' }] }] };
    await pageToPng(list([box]), { fonts: [{ family: '"Brand Sans"', bytes: new Uint8Array([1]) }] });
    expect(calls.fonts.length).toBeGreaterThan(0);
    expect(calls.fonts.every((font) => /bo-vsdx-png-\d+-0$/.test(font))).toBe(true);
  });

  test('rejects when the browser reports no face for a registered family', async () => {
    const { calls } = stubBrowser();
    (globalThis as unknown as { fonts: { load(): Promise<unknown[]> } }).fonts.load = () => Promise.resolve([]);
    await expect(pageToPng(list(), { fonts: [face] })).rejects.toMatchObject({ code: 'font-unavailable' });
    expect(calls.encoded).toBe(0);
  });

  test('rejects instead of painting a fallback when a face cannot load', async () => {
    const { calls } = stubBrowser({ failLoad: true });
    const error = await pageToPng(list(), { fonts: [face] }).catch((value: unknown) => value);
    expect(error).toBeInstanceOf(PngExportError);
    expect((error as PngExportError).code).toBe('font-unavailable');
    expect(calls.ctx).not.toContain('fillText');
    expect(calls.encoded).toBe(0);
  });

  test('rejects instead of painting a fallback when a document font fails to load', async () => {
    const { calls } = stubBrowser();
    (globalThis as unknown as { fonts: { load(): Promise<never> } }).fonts.load = () => Promise.reject(new Error('network'));
    await expect(pageToPng(list())).rejects.toMatchObject({ code: 'font-unavailable' });
    expect(calls.encoded).toBe(0);
  });

  test('rejects a page over the canvas budget', async () => {
    const { calls } = stubBrowser();
    await expect(pageToPng(list([]), { scale: 1000 })).rejects.toMatchObject({ code: 'canvas-too-large' });
    expect(calls.sizes).toEqual([]);
  });

  test('outlines an image that cannot decode, as the raster backend does', async () => {
    const { calls } = stubBrowser();
    (globalThis as { createImageBitmap?: unknown }).createImageBitmap = () => Promise.reject(new Error('bad image'));
    const image = { id: 'i', zOrder: 1, kind: 'image', assetId: 'media/a.emf', x: 0, y: 0, width: 10, height: 10 };
    const png = await pageToPng(list([image as never]), { media: () => new Uint8Array([1]) });
    expect(new TextDecoder().decode(png)).toBe('png');
    expect(calls.ctx).not.toContain('drawImage');
    expect(calls.ctx).toContain('strokeRect');
  });

  test('the on-screen painter still leaves an undecodable image blank', async () => {
    const { calls } = stubBrowser();
    const image = { id: 'i', zOrder: 1, kind: 'image', assetId: 'media/a.emf', x: 0, y: 0, width: 10, height: 10 };
    const canvas = new OffscreenCanvas(10, 10);
    await paintPage(canvas.getContext('2d') as unknown as CanvasRenderingContext2D, list([image as never]), 1, 1, { resolveImage: () => null });
    expect(calls.ctx).not.toContain('strokeRect');
  });

  test('rejects where no canvas exists', async () => {
    stubBrowser();
    (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = undefined;
    const error = await pageToPng(list([])).catch((value: unknown) => value) as PngExportError;
    expect(error.code).toBe('canvas-unavailable');
    expect(error.message).toContain('@betteroffice/vsdx-native');
  });

  test('stops without encoding when aborted', async () => {
    const { calls } = stubBrowser();
    const controller = new AbortController();
    controller.abort();
    await expect(pageToPng(list([]), { signal: controller.signal })).rejects.toBeDefined();
    expect(calls.encoded).toBe(0);
  });
});

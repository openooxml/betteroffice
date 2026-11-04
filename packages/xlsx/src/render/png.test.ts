import { afterEach, describe, expect, test } from 'bun:test';
import type { DisplayList } from '../display-list/types';
import { engineKind, renderPng, type XlsxPngEngine } from '../wasm/enginePng';
import { displayListToPng } from './png';

const text = { op: 'text', x: 4, y: 12, text: 'Total', fontSize: 11, color: '#000000', fontFamily: 'Brand Sans' } as const;
const list = (commands: DisplayList['commands'] = [text]): DisplayList => ({ width: 120.5, height: 40, commands });

interface Calls { sizes: [number, number][]; ctx: string[]; encoded: number; loads: string[]; transforms: number[][]; }

let undo: (() => void) | undefined;
afterEach(() => { undo?.(); undo = undefined; });

function stubBrowser(options: { loadFails?: boolean; canvas?: boolean } = {}) {
  const calls: Calls = { sizes: [], ctx: [], encoded: 0, loads: [], transforms: [] };
  class StubCanvas {
    constructor(width: number, height: number) { calls.sizes.push([width, height]); }
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'measureText') return (value: string) => ({ width: value.length });
          if (property === 'setTransform') return (...args: number[]) => { calls.ctx.push('setTransform'); calls.transforms.push(args); };
          if (property in target) return target[property as string];
          return () => { calls.ctx.push(String(property)); };
        },
        set(target, property, value) { target[property as string] = value; return true; },
      });
    }
    convertToBlob(request: { type: string }) { calls.encoded += 1; return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, fonts: scope.fonts };
  scope.OffscreenCanvas = options.canvas === false ? undefined : StubCanvas;
  scope.fonts = {
    load(font: string) {
      calls.loads.push(font);
      return options.loadFails ? Promise.reject(new Error('network')) : Promise.resolve([]);
    },
  };
  undo = () => { Object.assign(scope, saved); };
  return calls;
}

describe('displayListToPng', () => {
  test('waits for every face the editor would paint with before painting, then fills white behind the sheet', async () => {
    const calls = stubBrowser();
    const png = await displayListToPng(list([text, { ...text, bold: true }, text]), { dpr: 2, width: 241, height: 80 });
    expect(new TextDecoder().decode(png)).toBe('png');
    expect(calls.loads).toHaveLength(2);
    expect(calls.loads.every((font) => font.endsWith('px "Brand Sans", sans-serif'))).toBe(true);
    expect(calls.loads[1].startsWith('bold ')).toBe(true);
    expect(calls.sizes).toEqual([[241, 80]]);
    expect(calls.ctx.indexOf('fillText')).toBeGreaterThan(-1);
    expect(calls.ctx.lastIndexOf('fillRect')).toBeGreaterThan(calls.ctx.indexOf('fillText'));
    expect(calls.transforms[0]).toEqual([2, 0, 0, 2, 0, 0]);
  });

  test('rejects instead of painting a fallback when a face cannot load', async () => {
    const calls = stubBrowser({ loadFails: true });
    const error = await displayListToPng(list(), { dpr: 1, width: 121, height: 40 }).catch((value: unknown) => value);
    expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
    expect((error as Error).message).toContain('Brand Sans');
    expect(calls.sizes).toEqual([]);
    expect(calls.encoded).toBe(0);
  });

  test('rejects a PNG over the engine size limits before creating a canvas', async () => {
    const calls = stubBrowser();
    await expect(displayListToPng(list(), { dpr: 1, width: 16_385, height: 1 })).rejects.toMatchObject({ code: 'canvas-too-large' });
    await expect(displayListToPng(list(), { dpr: 1, width: 4097, height: 4097 })).rejects.toMatchObject({ code: 'canvas-too-large' });
    expect(calls.sizes).toEqual([]);
  });

  test('stops without encoding when aborted', async () => {
    const calls = stubBrowser();
    const controller = new AbortController();
    controller.abort();
    await expect(displayListToPng(list(), { dpr: 1, width: 121, height: 40, signal: controller.signal })).rejects.toBeDefined();
    expect(calls.encoded).toBe(0);
  });
});

describe.skipIf(engineKind === 'native')('the wasm PNG engine', () => {
  function engine(overrides: Partial<XlsxPngEngine> = {}) {
    const seen: { sheet: number; call: string; json: string }[] = [];
    let active = 0;
    const value: XlsxPngEngine = {
      onSheet(sheet, read) { const before = active; active = sheet; try { return read(); } finally { active = before; } },
      doc: {
        pngDisplayListJson(json) { seen.push({ sheet: active, call: 'viewport', json }); return JSON.stringify(list()); },
        rangePngDisplayListJson(json) { seen.push({ sheet: active, call: 'range', json }); return JSON.stringify({ ...list(), width: 60.2, height: 20.1 }); },
        renderPng(json) { seen.push({ sheet: active, call: 'raster viewport', json }); return new Uint8Array([1]); },
        renderRangePng(json) { seen.push({ sheet: active, call: 'raster range', json }); return new Uint8Array([2]); },
      },
      rasterAvailable: false,
      ...overrides,
    };
    return { value, seen };
  }

  test('paints a viewport through the canvas at the requested scale, capped like the engine', async () => {
    const calls = stubBrowser();
    const { value, seen } = engine();
    await renderPng(value, { sheet: 1, viewport: { x: 0, y: 0, width: 300, height: 120 } }, { scale: 2, maxWidth: 101 });
    expect(seen.map(({ sheet, call }) => [sheet, call])).toEqual([[1, 'viewport']]);
    expect(JSON.parse(seen[0].json).width).toBeLessThan(51);
    expect(calls.sizes[0][0]).toBeLessThanOrEqual(101);
    expect(calls.sizes[0][1]).toBe(240);
    expect(calls.transforms[0]).toEqual([2, 0, 0, 2, 0, 0]);
  });

  test('paints a range from the engine-scaled list at its own size', async () => {
    const calls = stubBrowser();
    const { value, seen } = engine();
    await renderPng(value, { sheet: 0, range: 'A1:B2' }, { scale: 1.5, maxHeight: 30 });
    expect(JSON.parse(seen[0].json)).toEqual({ range: 'A1:B2', scale: 1.5, maxHeight: 30 });
    expect(calls.sizes).toEqual([[61, 21]]);
    expect(calls.transforms[0]).toEqual([1, 0, 0, 1, 0, 0]);
  });

  test("draws the sheet's used range when the range is 'used'", async () => {
    stubBrowser();
    const { value, seen } = engine();
    await renderPng(value, { sheet: 1, range: 'used' }, { scale: 2 });
    expect(seen.map(({ sheet, call }) => [sheet, call])).toEqual([[1, 'range']]);
    expect(JSON.parse(seen[0].json)).toEqual({ scale: 2 });
  });

  test('without a canvas, uses the rasterizer the build carries, then a session rasterizer, then names -native', async () => {
    stubBrowser({ canvas: false });
    const viewport = { x: 0, y: 0, width: 10, height: 10 };
    const raster = engine({ rasterAvailable: true });
    expect([...await renderPng(raster.value, { sheet: 0, range: 'A1' }, {})]).toEqual([2]);
    const rasterized: string[] = [];
    const lazy = engine({ rasterize: async (json) => { rasterized.push(json); return new Uint8Array([3]); } });
    expect([...await renderPng(lazy.value, { sheet: 0, viewport }, {})]).toEqual([3]);
    expect(lazy.seen.map(({ call }) => call)).toEqual(['viewport']);
    expect(rasterized).toHaveLength(1);
    await expect(renderPng(engine().value, { sheet: 0, viewport }, {})).rejects.toMatchObject({
      name: 'PngExportError', code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/xlsx-native'),
    });
  });
});

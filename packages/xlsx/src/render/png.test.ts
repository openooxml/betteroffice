import { afterEach, describe, expect, test } from 'bun:test';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { DisplayList } from '../display-list/types';
import { engineKind, renderPng, type XlsxPngEngine } from '../wasm/enginePng';
import { displayListToPng } from './png';

const text = { op: 'text', x: 4, y: 12, text: 'Total', fontSize: 11, color: '#000000', fontFamily: 'Brand Sans' } as const;
const list = (commands: DisplayList['commands'] = [text]): DisplayList => ({ width: 120.5, height: 40, commands });

interface Calls { sizes: [number, number][]; ctx: string[]; encoded: number; loads: string[]; transforms: number[][]; painted: string[]; faces: { family: string; weight: string; style: string }[]; held: unknown[]; }

let undo: (() => void) | undefined;
afterEach(() => { undo?.(); undo = undefined; });

function stubBrowser(options: { loadFails?: boolean; canvas?: boolean; faceFails?: boolean } = {}) {
  const calls: Calls = { sizes: [], ctx: [], encoded: 0, loads: [], transforms: [], painted: [], faces: [], held: [] };
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
        set(target, property, value) { if (property === 'font') calls.painted.push(String(value)); target[property as string] = value; return true; },
      });
    }
    convertToBlob(request: { type: string }) { calls.encoded += 1; return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  const scope = globalThis as Record<string, unknown>;
  class StubFace {
    family: string;
    constructor(family: string, _bytes: ArrayBuffer, descriptors: { weight: string; style: string }) {
      this.family = family;
      calls.faces.push({ family, ...descriptors });
    }
    load() { return options.faceFails ? Promise.reject(new Error('bad font')) : Promise.resolve(this); }
  }
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, document: scope.document, fonts: scope.fonts, FontFace: scope.FontFace };
  scope.OffscreenCanvas = options.canvas === false ? undefined : StubCanvas;
  if (options.canvas === false) scope.document = undefined;
  scope.FontFace = StubFace;
  scope.fonts = {
    add(face: unknown) { calls.held.push(face); },
    delete(face: unknown) { calls.held.splice(calls.held.indexOf(face), 1); return true; },
    load(font: string) {
      calls.loads.push(font);
      if (options.loadFails) return Promise.reject(new Error('network'));
      return Promise.resolve(calls.held.filter((face) => font.includes(`"${(face as StubFace).family}"`)));
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

describe('supplied faces', () => {
  const calibri = { op: 'text', x: 4, y: 12, text: 'Total', fontSize: 11, color: '#000000', fontFamily: 'Calibri' } as const;

  test('paint cell text under a private family, matched by family and style, and are dropped afterwards', async () => {
    const calls = stubBrowser();
    const regular = { family: 'calibri', bytes: new Uint8Array([1]) };
    await displayListToPng(list([calibri, { ...calibri, bold: true }, { ...text, fontFamily: 'Cambria' }]), { dpr: 1, width: 121, height: 40, fonts: [regular] });
    expect(calls.faces).toEqual([{ family: expect.stringMatching(/^bo-xlsx-png-\d+-0$/), weight: '400', style: 'normal' }]);
    const fills = calls.painted.filter((font) => font.includes('px'));
    expect(fills.filter((font) => /"bo-xlsx-png-\d+-0", sans-serif$/.test(font))).toHaveLength(2);
    expect(fills.some((font) => font.endsWith('"Cambria", sans-serif'))).toBe(true);
    expect(calls.held).toEqual([]);
  });

  test('a supplied face that cannot load rejects instead of painting a substitute', async () => {
    const calls = stubBrowser({ faceFails: true });
    await expect(displayListToPng(list([calibri]), { dpr: 1, width: 121, height: 40, fonts: [{ family: 'Calibri', bytes: new Uint8Array([1]) }] }))
      .rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable', message: expect.stringContaining('Calibri') });
    expect(calls.encoded).toBe(0);
  });

  test('chart text heading its font list with a supplied family paints with that face, as the editor page does', async () => {
    const calls = stubBrowser();
    const chart = { ...calibri, fontFamily: 'Calibri, sans-serif', chart: true };
    const generic = { ...calibri, fontFamily: 'sans-serif', chart: true };
    await displayListToPng(list([chart as never, generic as never]), { dpr: 1, width: 121, height: 40, fonts: [{ family: 'Calibri', bytes: new Uint8Array([1]) }] });
    const fills = calls.painted.filter((font) => font.includes('px'));
    expect(fills.some((font) => /"bo-xlsx-png-\d+-0", Calibri, sans-serif$/.test(font))).toBe(true);
    expect(fills.some((font) => font.endsWith('px sans-serif'))).toBe(true);
    expect(calls.faces).toHaveLength(1);
  });

  test('a bold run takes the bold face when one is supplied', async () => {
    const calls = stubBrowser();
    await displayListToPng(list([{ ...calibri, bold: true }]), { dpr: 1, width: 121, height: 40, fonts: [
      { family: 'Calibri', bytes: new Uint8Array([1]) }, { family: 'Calibri', bold: true, bytes: new Uint8Array([2]) },
    ] });
    expect(calls.faces.map((face) => face.weight)).toEqual(['700']);
  });
});

describe.skipIf(engineKind === 'native')('the wasm PNG engine', () => {
  function engine() {
    const seen: { sheet: number; call: string; json: string }[] = [];
    let active = 0;
    const value: XlsxPngEngine = {
      onSheet(sheet, read) { const before = active; active = sheet; try { return read(); } finally { active = before; } },
      doc: {
        pngDisplayListJson(json) { seen.push({ sheet: active, call: 'viewport', json }); return JSON.stringify(list()); },
        rangePngDisplayListJson(json) { seen.push({ sheet: active, call: 'range', json }); return JSON.stringify({ ...list(), width: 60.2, height: 20.1 }); },
      },
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

  test('without a canvas, viewport and range exports reject and name the native package', async () => {
    stubBrowser({ canvas: false });
    const viewport = { x: 0, y: 0, width: 10, height: 10 };
    const { value, seen } = engine();
    for (const target of [{ sheet: 0, viewport }, { sheet: 0, range: 'A1' }]) {
      const png = renderPng(value, target, {});
      await expect(png).rejects.toBeInstanceOf(PngExportError);
      await expect(png).rejects.toMatchObject({
        name: 'PngExportError', code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/xlsx-native'),
      });
    }
    expect(seen).toEqual([]);
  });
});

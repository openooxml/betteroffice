import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import * as bundled from '@betteroffice/fonts';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { DisplayList } from '../display-list/types';
import { engineKind, renderPng, type XlsxPngEngine } from '../wasm/enginePng';
import { displayListToPng } from './png';

const text = { op: 'text', x: 4, y: 12, text: 'Total', fontSize: 11, color: '#000000', fontFamily: 'Brand Sans' } as const;
const list = (commands: DisplayList['commands'] = [text]): DisplayList => ({ width: 120.5, height: 40, commands });

interface Calls { sizes: [number, number][]; ctx: string[]; encoded: number; loads: string[]; transforms: number[][]; painted: string[]; faces: { family: string; weight: string; style: string }[]; held: unknown[]; bytes: number[]; }

let undo: (() => void) | undefined;
afterEach(() => { undo?.(); undo = undefined; });

function stubBrowser(options: { loadFails?: boolean; canvas?: boolean; faceFails?: boolean } = {}) {
  const calls: Calls = { sizes: [], ctx: [], encoded: 0, loads: [], transforms: [], painted: [], faces: [], held: [], bytes: [] };
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
    constructor(family: string, bytes: ArrayBuffer | Uint8Array, descriptors: { weight: string; style: string }) {
      this.family = family;
      calls.bytes.push(bytes.byteLength);
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

  test('without fonts, a face the native engine would draw with fails to load and the export rejects naming it', async () => {
    stubBrowser();
    const failing = spyOn(bundled, 'loadBundledFontBytes').mockImplementation(() => Promise.reject(new Error('offline')));
    try {
      const georgia: XlsxPngEngine = {
        onSheet: (_sheet, read) => read(),
        doc: {
          pngDisplayListJson: () => JSON.stringify(list([{ ...text, fontFamily: 'Georgia' }])),
          rangePngDisplayListJson: () => JSON.stringify(list([{ ...text, fontFamily: 'Georgia' }])),
        },
      };
      await expect(renderPng(georgia, { sheet: 0, viewport: { x: 0, y: 0, width: 100, height: 40 } }, {})).rejects.toMatchObject({
        name: 'PngExportError', code: 'font-unavailable', message: expect.stringContaining('for Georgia could not load: offline'),
      });
      expect(failing).toHaveBeenCalled();
    } finally { failing.mockRestore(); }
  });

  test('without fonts, the sheet paints with the bundled face the native engine draws its font with', async () => {
    const calls = stubBrowser();
    const { value } = engine();
    await renderPng(value, { sheet: 0, viewport: { x: 0, y: 0, width: 100, height: 40 } }, {});
    const face = bundled.resolveGlyphFallbackFace('Brand Sans', false, false, 'word') ?? bundled.resolveLastResortChoice('Brand Sans', false, false, 'word').face;
    expect(calls.bytes).toEqual([(await bundled.loadBundledFontBytes(face)).byteLength]);
  });

  const sheetOf = (fontFamily: string, bold = false): XlsxPngEngine => ({
    onSheet: (_sheet, read) => read(),
    doc: {
      pngDisplayListJson: () => JSON.stringify(list([{ ...text, fontFamily, bold }])),
      rangePngDisplayListJson: () => JSON.stringify(list([{ ...text, fontFamily, bold }])),
    },
  });
  const box = { sheet: 0, viewport: { x: 0, y: 0, width: 100, height: 40 } };

  const nativeFace = (family: string, bold: boolean) => bundled.resolveBundledFamilyFace(family, bold, false)
    ?? bundled.resolveGlyphFallbackFace(family, bold, false, 'word') ?? bundled.resolveLastResortChoice(family, bold, false, 'word').face;

  test('without fonts, a bold run on a regular face installs it regular so the bold is synthesized as natively', async () => {
    const calls = stubBrowser();
    await renderPng(sheetOf('Calibri Light', true), box, {});
    expect(nativeFace('Calibri Light', true).weight).toBe(400);
    expect(calls.faces.map((installed) => installed.weight)).toEqual(['400']);
    expect(calls.painted.some((font) => font.startsWith('bold ') && font.includes('bo-xlsx-png'))).toBe(true);
  });

  test('without fonts, a regular run on a bold face paints with that face, installed bold', async () => {
    const calls = stubBrowser();
    await renderPng(sheetOf('Arial Black'), box, {});
    expect(nativeFace('Arial Black', false).weight).toBe(700);
    expect(calls.faces.map((installed) => installed.weight)).toEqual(['700']);
    expect(calls.painted.some((font) => !font.startsWith('bold ') && font.includes('bo-xlsx-png'))).toBe(true);
  });

  test('without fonts, a family whose bundled face is in a package not installed paints with the next face natively', async () => {
    const calls = stubBrowser();
    const primary = bundled.resolveBundledFamilyFace('Calibri', false, false)!;
    const original = bundled.bundledFaceInstalled;
    const installed = spyOn(bundled, 'bundledFaceInstalled')
      .mockImplementation((face) => face.file === primary.file ? Promise.resolve(false) : original(face));
    try {
      await renderPng(sheetOf('Calibri'), box, {});
      const next = bundled.resolveGlyphFallbackFace('Calibri', false, false, 'word') ?? bundled.resolveLastResortChoice('Calibri', false, false, 'word').face;
      expect(next.file).not.toBe(primary.file);
      expect(calls.bytes).toEqual([(await bundled.loadBundledFontBytes(next)).byteLength]);
    } finally { installed.mockRestore(); }
  });

  test('without fonts, faces load strictly, so a missing add-on rejects instead of painting another face', async () => {
    stubBrowser();
    const loading = spyOn(bundled, 'loadBundledFontBytes');
    try {
      await renderPng(sheetOf('Times New Roman', true), box, {});
      expect(loading.mock.calls.length).toBeGreaterThan(0);
      expect(loading.mock.calls.every((call) => call[2]?.strict === true)).toBe(true);
    } finally { loading.mockRestore(); }
  });

  test('without fonts, an aborted export rejects with the abort reason and starts no load', async () => {
    stubBrowser();
    const loading = spyOn(bundled, 'loadBundledFontBytes');
    try {
      const controller = new AbortController();
      controller.abort(new Error('stopped'));
      await expect(renderPng(sheetOf('Cambria'), box, { signal: controller.signal })).rejects.toThrow('stopped');
      expect(loading).not.toHaveBeenCalled();
    } finally { loading.mockRestore(); }
  });

  test('without fonts, an export aborted while a face loads rejects at once', async () => {
    stubBrowser();
    const stalled = spyOn(bundled, 'loadBundledFontBytes').mockImplementation(() => new Promise(() => {}));
    try {
      const controller = new AbortController();
      const exported = renderPng(sheetOf('Courier New'), box, { signal: controller.signal });
      await new Promise((settled) => setTimeout(settled, 10));
      controller.abort(new Error('stopped'));
      await expect(exported).rejects.toThrow('stopped');
      expect(stalled).toHaveBeenCalled();
    } finally { stalled.mockRestore(); }
  });

  const cells = (commands: Array<{ fontFamily: string; text: string; bold?: boolean }>): XlsxPngEngine => {
    const json = JSON.stringify(list(commands.map((cell, index) => ({ ...text, y: 12 + index * 16, ...cell }))));
    return { onSheet: (_sheet, read) => read(), doc: { pngDisplayListJson: () => json, rangePngDisplayListJson: () => json } };
  };
  const korean = bundled.resolveScriptFallbackFace('cjk-kr', false, false)!;
  const KOREAN_BYTES = 7777;
  const sameFace = (a: bundled.BundledFontFace, b: bundled.BundledFontFace) => a.family === b.family && a.weight === b.weight && a.style === b.style;
  async function nativeChain(family: string, scripts: bundled.BundledFontScript[][]): Promise<bundled.BundledFontFace[]> {
    const chain: bundled.BundledFontFace[] = [];
    const push = async (face: bundled.BundledFontFace | undefined) => {
      if (face && await bundled.bundledFaceInstalled(face) && !chain.some((known) => sameFace(known, face))) chain.push(face);
    };
    for (const request of scripts) {
      const primary = bundled.resolveBundledFamilyFace(family, false, false);
      const installed = primary && await bundled.bundledFaceInstalled(primary) ? primary : undefined;
      await push(installed);
      for (const script of request) await push(bundled.resolveScriptFallbackFace(script, false, false, family));
      await push(bundled.resolveGlyphFallbackFace(family, false, false, 'word'));
      const last = bundled.resolveLastResortChoice(family, false, false, 'word');
      if (!last.onlyIfPrimaryFails || !installed) await push(last.face);
    }
    return chain;
  }
  function koreanLoads(stalled?: bundled.BundledFontFace) {
    const original = bundled.loadBundledFontBytes;
    const files: string[] = [];
    const spy = spyOn(bundled, 'loadBundledFontBytes').mockImplementation((face, ...rest) => {
      files.push(face.file);
      if (face.file === stalled?.file) return new Promise(() => {});
      return face.file === korean.file ? Promise.resolve(new ArrayBuffer(KOREAN_BYTES)) : original(face, ...rest);
    });
    return { files, restore: () => spy.mockRestore() };
  }

  test('without fonts, a sheet with Latin text only installs one face per font and loads no script face', async () => {
    const calls = stubBrowser();
    const loads = koreanLoads();
    try {
      await renderPng(cells([
        { fontFamily: 'Calibri', text: 'Total' }, { fontFamily: 'Calibri', text: 'Überschrift – 2024', bold: true },
        { fontFamily: 'Cambria', text: '€ 12,50 ✓' },
      ]), box, {});
      expect(calls.faces.length).toBe(3);
      expect(loads.files).not.toContain(korean.file);
      expect(calls.painted.filter((font) => font.includes('bo-xlsx-png')).every((font) => font.split('bo-xlsx-png').length === 2)).toBe(true);
    } finally { loads.restore(); }
  });

  test('without fonts, a font with Korean text anywhere on the sheet paints all its text through the native chain', async () => {
    const calls = stubBrowser();
    const loads = koreanLoads();
    try {
      await renderPng(cells([{ fontFamily: 'Calibri', text: 'Total' }, { fontFamily: 'Calibri', text: '합계 2024' }]), box, {});
      const chain = await nativeChain('Calibri', [[], ['cjk-kr']]);
      expect(chain.some((face) => sameFace(face, korean))).toBe(true);
      expect(calls.faces.length).toBe(chain.length);
      expect(calls.bytes).toContain(KOREAN_BYTES);
      const listed = calls.painted.filter((font) => font.includes('bo-xlsx-png'));
      expect(listed.length).toBeGreaterThan(0);
      for (const font of listed) {
        const names = [...font.matchAll(/"(bo-xlsx-png-[^"]+)"/g)].map((match) => match[1]);
        expect(names.length).toBe(chain.length);
        expect(font.endsWith(', sans-serif')).toBe(true);
      }
    } finally { loads.restore(); }
  });

  test('without fonts, a family with no bundled face draws Korean text with the script face first, as natively', async () => {
    const calls = stubBrowser();
    const loads = koreanLoads();
    try {
      await renderPng(cells([{ fontFamily: 'Brand Gothic', text: '합계 2024' }]), box, {});
      const chain = await nativeChain('Brand Gothic', [['cjk-kr']]);
      expect(sameFace(chain[0]!, korean)).toBe(true);
      expect(calls.bytes[0]).toBe(KOREAN_BYTES);
      expect(calls.faces.length).toBe(chain.length);
    } finally { loads.restore(); }
  });

  test('without fonts, a script face whose package is not installed is left out and never loaded', async () => {
    const calls = stubBrowser();
    const loads = koreanLoads();
    const original = bundled.bundledFaceInstalled;
    const installed = spyOn(bundled, 'bundledFaceInstalled')
      .mockImplementation((face) => face.file === korean.file ? Promise.resolve(false) : original(face));
    try {
      await renderPng(cells([{ fontFamily: 'Calibri', text: '합계 2024' }]), box, {});
      expect(loads.files).not.toContain(korean.file);
      expect(calls.bytes).not.toContain(KOREAN_BYTES);
      expect(calls.faces.length).toBe((await nativeChain('Calibri', [['cjk-kr']])).length);
    } finally { installed.mockRestore(); loads.restore(); }
  });

  test('without fonts, an export aborted while a script face loads rejects at once', async () => {
    stubBrowser();
    const thai = bundled.resolveScriptFallbackFace('thai', false, false)!;
    const loads = koreanLoads(thai);
    try {
      const controller = new AbortController();
      const exported = renderPng(cells([{ fontFamily: 'Brand Gothic', text: 'สวัสดี' }]), box, { signal: controller.signal });
      await new Promise((settled) => setTimeout(settled, 10));
      controller.abort(new Error('stopped'));
      await expect(exported).rejects.toThrow('stopped');
      expect(loads.files).toContain(thai.file);
    } finally { loads.restore(); }
  });

  test('with fonts, the sheet paints only with the faces passed', async () => {
    const calls = stubBrowser();
    const { value } = engine();
    await renderPng(value, { sheet: 0, viewport: { x: 0, y: 0, width: 100, height: 40 } }, { fonts: [{ family: 'Brand Sans', bytes: new Uint8Array([1, 2, 3]) }] });
    expect(calls.bytes).toEqual([3]);
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

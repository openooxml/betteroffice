import { afterEach, beforeAll, describe, expect, test, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as fonts from '@betteroffice/fonts';
import { initWasm, openPresentation } from '../index';
import { engineKind } from './engine';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;
let font: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx, face] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = new Uint8Array(pptx);
  font = new Uint8Array(face);
});

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stubBrowser() {
  const sizes: [number, number][] = [];
  const sources: Uint8Array[] = [];
  const texts: string[] = [];
  class Canvas {
    constructor(width: number, height: number) { sizes.push([width, height]); }
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'createLinearGradient' || property === 'createRadialGradient' || property === 'createPattern') return () => ({ addColorStop: () => undefined });
          if (property === 'measureText') return (text: string) => ({ width: text.length });
          if (property === 'fillText') return (text: string) => { texts.push(text); };
          return property in target ? target[property as string] : () => undefined;
        },
        set(target, property, value) { target[property as string] = value; return true; },
      });
    }
    convertToBlob(request: { type: string }) { return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  class Face {
    family: string;
    constructor(family: string, source: ArrayBuffer) { this.family = family; sources.push(new Uint8Array(source)); }
    load() { return Promise.resolve(this); }
  }
  const held: Face[] = [];
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, FontFace: scope.FontFace, fonts: scope.fonts, createImageBitmap: scope.createImageBitmap, Image: scope.Image };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, FontFace: Face, Image: undefined,
    createImageBitmap: () => Promise.reject(new Error('no decoder in this test')),
    fonts: {
      add(face: Face) { held.push(face); },
      delete(face: Face) { held.splice(held.indexOf(face), 1); return true; },
      load(request: string) { return Promise.resolve(held.filter((face) => request.includes(`"${face.family}"`))); },
    },
  });
  restore = () => { Object.assign(scope, saved); };
  return { sizes, sources, held, texts };
}

const native = engineKind === 'native';
const isPng = (bytes: Uint8Array) => bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;
const pngSize = (bytes: Uint8Array): [number, number] => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return [view.getUint32(16), view.getUint32(20)];
};

describe('PresentationHandle.renderPng', () => {
  test('draws a slide by index or ID without moving the slide hit testing reads', async () => {
    const browser = stubBrowser();
    const handle = openPresentation(fixture, { clientId: 7101, fonts: [{ family: 'Liberation Sans', bytes: font }] });
    try {
      const shown = handle.layoutSlide(0);
      const point = { x: shown.width / 2, y: shown.height / 2 };
      const before = handle.hitTest(point.x, point.y);
      const scaled = await handle.renderPng({ slide: 1 }, { scale: 2 });
      expect(handle.hitTest(point.x, point.y)).toEqual(before);
      const id = handle.snapshot().slides[1].id;
      const plain = await handle.renderPng({ slide: id });
      const second = handle.layoutSlide(1);
      const expected: [number, number][] = [[Math.ceil(second.width * 2), Math.ceil(second.height * 2)], [Math.ceil(second.width), Math.ceil(second.height)]];
      if (native) {
        expect([scaled, plain].every(isPng)).toBe(true);
        expect([pngSize(scaled), pngSize(plain)]).toEqual(expected);
      } else {
        expect(new TextDecoder().decode(scaled)).toBe('png');
        expect(browser.sizes).toEqual(expected);
      }
      expect(browser.held).toEqual([]);
    } finally { handle.dispose(); }
  });

  test('refuses targets the native renderer refuses, with the same messages', async () => {
    stubBrowser();
    const handle = openPresentation(fixture, { clientId: 7102 });
    try {
      const count = handle.snapshot().slides.length;
      await expect(handle.renderPng({ slide: count })).rejects.toThrow(`slide index ${count} out of range for ${count} slide(s)`);
      await expect(handle.renderPng({ slide: 'missing' })).rejects.toThrow('no slide with ID "missing"');
      await expect(handle.renderPng({ slide: 1.5 })).rejects.toThrow('slide must be an ID (string) or an index (number)');
      await expect(handle.renderPng({ slide: 0, extra: 1 } as never)).rejects.toThrow('target must be { slide } with a slide index or ID');
      await expect(handle.renderPng({ slide: 0 }, { background: 'red' as never })).rejects.toThrow('background must be');
      await expect(handle.renderPng({ slide: 0 }, { scale: 0 })).rejects.toThrow('scale must be finite and positive');
      await expect(handle.renderPng({ slide: 0 }, { scale: 100 })).rejects.toMatchObject({ name: 'PngExportError', code: 'canvas-too-large' });
      await expect(handle.renderPng({ slide: 0 }, { maxShadowPixels: -1 })).rejects.toThrow('invalid shadow pixel budget');
    } finally { handle.dispose(); }
  });

  test('a presentation disposed during an export rejects it instead of returning a slide without its pictures', async () => {
    const browser = stubBrowser();
    const handle = openPresentation(fixture, { clientId: 7104, fonts: [{ family: 'Liberation Sans', bytes: font }] });
    const closed: boolean[] = [];
    let decodes = 0;
    (globalThis as Record<string, unknown>).createImageBitmap = () => {
      decodes++;
      handle.dispose();
      const bitmap = { width: 4, height: 4, close: () => { closed.push(true); } };
      return Promise.resolve(bitmap);
    };
    const exported = handle.renderPng({ slide: 0 });
    if (native) handle.dispose();
    await expect(exported).rejects.toThrow('presentation handle is disposed');
    await expect(handle.renderPng({ slide: 0 })).rejects.toThrow('presentation handle is disposed');
    if (!native) {
      expect(decodes).toBe(1);
      expect(closed).toEqual([true]);
    }
    expect(browser.held).toEqual([]);
  });

  test('a slide with metafile pictures the canvas cannot draw rejects instead of leaving them out', async () => {
    const browser = stubBrowser();
    const deck = new Uint8Array(await readFile(resolve(root, 'crates/pptx-render/tests/fixtures/metafile-svg-pictures.pptx')));
    const handle = openPresentation(deck, { clientId: 7105, fonts: [{ family: 'Liberation Sans', bytes: font }] });
    try {
      const exported = handle.renderPng({ slide: 0 });
      if (native) expect(isPng(await exported)).toBe(true);
      else {
        await expect(exported).rejects.toMatchObject({ name: 'PngExportError', code: 'canvas-unavailable' });
        expect(browser.sizes).toEqual([]);
      }
    } finally { handle.dispose(); }
  });

  test('paints each run with the face layout measured it with, as registered, not as the caller changed it afterwards', async () => {
    const browser = stubBrowser();
    const opened = { family: 'Liberation Sans', bytes: font.slice() };
    const handle = openPresentation(fixture, { clientId: 7103, fonts: [opened] });
    try {
      const slide = handle.snapshot().slides[0]!;
      handle.addTextBox(slide.id, { name: 'Registered', rect: { x: 100000, y: 100000, width: 3000000, height: 800000 }, text: 'Registered face', style: { fontFamily: 'Liberation Sans', fontSizePt: 18 } });
      const later = { family: 'Liberation Sans', bold: true, bytes: font.slice() };
      handle.registerFont(later);
      for (const face of [opened, later]) { face.bytes.fill(0); face.family = 'Changed'; }
      const png = await handle.renderPng({ slide: 0 });
      if (native) expect(isPng(png)).toBe(true);
      else {
        expect(browser.sources.some((source) => source.length === font.length && source.every((byte, index) => byte === font[index]))).toBe(true);
        for (const source of browser.sources) expect(source.some((byte) => byte !== 0)).toBe(true);
      }
    } finally { handle.dispose(); }
  });
});

test.skipIf(native)('canvas export lays the slide out with the bundled faces when no fonts are passed', async () => {
  const browser = stubBrowser();
  const handle = openPresentation(fixture, { clientId: 7195 });
  try {
    const before = handle.layoutSlide(0);
    await handle.renderPng({ slide: 0 });
    expect(browser.sources.length).toBeGreaterThan(0);
    expect(handle.layoutSlide(0)).toEqual(before);
    const first = browser.sources.length;
    await handle.renderPng({ slide: 0 });
    expect(browser.sources.length).toBe(2 * first);
  } finally { handle.dispose(); }
});

test.skipIf(native)('canvas export draws the slide it was asked for when the deck changes while faces load', async () => {
  const browser = stubBrowser();
  const handle = openPresentation(fixture, { clientId: 7193 });
  try {
    const slides = handle.snapshot().slides;
    const asked = slides[1]!.id;
    handle.addTextBox(asked, { name: 'Marker', rect: { x: 100000, y: 100000, width: 3000000, height: 800000 }, text: 'Asked for', style: { fontFamily: 'Calibri', fontSizePt: 18 } });
    const moved = handle.renderPng({ slide: 1 });
    handle.moveSlide(asked, 0);
    await moved;
    expect(browser.texts.join(' ')).toContain('Asked for');
    const removed = handle.renderPng({ slide: asked });
    handle.deleteSlide(asked);
    await expect(removed).rejects.toThrow(`no slide with ID ${JSON.stringify(asked)}`);
  } finally { handle.dispose(); }
});

test.skipIf(native)('canvas export loads the faces of text added while faces load', async () => {
  const browser = stubBrowser();
  const box = (handle: ReturnType<typeof openPresentation>, slideId: string, name: string, style: { fontFamily: string; bold?: boolean; italic?: boolean }, y: number) =>
    handle.addTextBox(slideId, { name, rect: { x: 100000, y, width: 3000000, height: 800000 }, text: name, style: { ...style, fontSizePt: 18 } });
  const first = { fontFamily: 'Courier New', bold: true, italic: true };
  const later = { fontFamily: 'Georgia', bold: true, italic: true };
  const edited = openPresentation(fixture, { clientId: 7192 });
  const fresh = openPresentation(fixture, { clientId: 7191 });
  const original = fonts.loadBundledFontBytes;
  let edits = 0;
  const editedSlide = edited.snapshot().slides[0]!.id;
  const loading = spyOn(fonts, 'loadBundledFontBytes').mockImplementation((...args: Parameters<typeof original>) => {
    if (edits === 0) {
      edits += 1;
      box(edited, editedSlide, 'Added later', later, 1000000);
    }
    return original(...args);
  });
  try {
    box(edited, editedSlide, 'First', first, 100000);
    await edited.renderPng({ slide: 0 });
    loading.mockRestore();
    expect(edits).toBe(1);
    const whileLoading = browser.sources.map((source) => source.length).sort();
    browser.sources.length = 0;
    const freshSlide = fresh.snapshot().slides[0]!.id;
    box(fresh, freshSlide, 'First', first, 100000);
    box(fresh, freshSlide, 'Added later', later, 1000000);
    await fresh.renderPng({ slide: 0 });
    expect(whileLoading).toEqual(browser.sources.map((source) => source.length).sort());
  } finally { loading.mockRestore(); edited.dispose(); fresh.dispose(); }
});

test.skipIf(native)('a presentation disposed before its first canvas export starts rejects it', async () => {
  stubBrowser();
  const handle = openPresentation(fixture, { clientId: 7190 });
  const exported = handle.renderPng({ slide: 0 });
  handle.dispose();
  await expect(exported).rejects.toThrow('presentation handle is disposed');
});

test.skipIf(native)('disposing a presentation frees its canvas export renderer', async () => {
  stubBrowser();
  const { PptxRenderer } = await import('./engine');
  const free = spyOn(PptxRenderer.prototype, 'free');
  try {
    const handle = openPresentation(fixture, { clientId: 7194 });
    await handle.renderPng({ slide: 0 });
    free.mockClear();
    handle.dispose();
    expect(free).toHaveBeenCalledTimes(2);
  } finally { free.mockRestore(); }
});

test.skipIf(!native)('native export paints slide text from the font packs with no fonts passed', async () => {
  const handle = openPresentation(fixture, { clientId: 7199 });
  try {
    const before = handle.layoutSlide(0);
    const png = await handle.renderPng({ slide: 0 });
    expect(isPng(png)).toBe(true);
    expect(handle.layoutSlide(0)).toEqual(before);
    expect(await handle.renderPng({ slide: 0 })).toEqual(png);
  } finally { handle.dispose(); }
});

test.skipIf(!native)('native export keeps a registered fallback face a fallback', async () => {
  const handle = openPresentation(fixture, { clientId: 7198, fallbackFonts: [{ family: 'Liberation Sans', bytes: font }] });
  try {
    expect(isPng(await handle.renderPng({ slide: 0 }))).toBe(true);
    handle.registerFont({ family: 'Liberation Sans', bytes: font });
    expect(isPng(await handle.renderPng({ slide: 0 }, { fonts: [{ family: 'Liberation Sans', bold: true, bytes: font }] }))).toBe(true);
  } finally { handle.dispose(); }
});

test.skipIf(!native)('native export keeps the face layout measured when options.fonts names it again', async () => {
  const other = new Uint8Array(await Bun.file(new URL('../../../fonts/assets/ComicRelief-Regular.ttf', import.meta.url)).arrayBuffer());
  const handle = openPresentation(fixture, { clientId: 7197, fonts: [{ family: 'Liberation Sans', bytes: font }] });
  try {
    const slide = handle.snapshot().slides[0]!;
    handle.addTextBox(slide.id, { name: 'Measured', rect: { x: 100000, y: 100000, width: 3000000, height: 800000 }, text: 'Measured face', style: { fontFamily: 'Liberation Sans', fontSizePt: 18 } });
    handle.addTextBox(slide.id, { name: 'Measured bold', rect: { x: 100000, y: 1000000, width: 3000000, height: 800000 }, text: 'Measured bold', style: { fontFamily: 'Liberation Sans', fontSizePt: 18, bold: true } });
    const plain = await handle.renderPng({ slide: 0 });
    expect(await handle.renderPng({ slide: 0 }, { fonts: [{ family: 'Liberation Sans', bytes: other }, { family: 'Liberation Sans', bold: true, bytes: other }] })).toEqual(plain);
  } finally { handle.dispose(); }
});

test.skipIf(!native)('disposing a presentation frees its export renderer', async () => {
  const binding = require('../../../../bindings/pptx-native/index.js') as { PptxRenderer: { prototype: { free(): void } } };
  const free = spyOn(binding.PptxRenderer.prototype, 'free');
  try {
    const handle = openPresentation(fixture, { clientId: 7196 });
    await handle.renderPng({ slide: 0 });
    free.mockClear();
    handle.dispose();
    expect(free).toHaveBeenCalledTimes(2);
  } finally { free.mockRestore(); }
});

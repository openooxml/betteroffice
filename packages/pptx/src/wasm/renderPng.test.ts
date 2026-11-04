import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
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
  class Canvas {
    constructor(width: number, height: number) { sizes.push([width, height]); }
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'createLinearGradient' || property === 'createRadialGradient' || property === 'createPattern') return () => ({ addColorStop: () => undefined });
          if (property === 'measureText') return (text: string) => ({ width: text.length });
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
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, FontFace: scope.FontFace, fonts: scope.fonts, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, FontFace: Face,
    createImageBitmap: () => Promise.reject(new Error('no decoder in this test')),
    fonts: {
      add(face: Face) { held.push(face); },
      delete(face: Face) { held.splice(held.indexOf(face), 1); return true; },
      load(request: string) { return Promise.resolve(held.filter((face) => request.includes(`"${face.family}"`))); },
    },
  });
  restore = () => { Object.assign(scope, saved); };
  return { sizes, sources, held };
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

  test('paints each run with the face layout measured it with, as registered, not as the caller changed it afterwards', async () => {
    const browser = stubBrowser();
    const opened = { family: 'Liberation Sans', bytes: font.slice() };
    const handle = openPresentation(fixture, { clientId: 7103, fonts: [opened] });
    try {
      const later = { family: 'Liberation Sans', bold: true, bytes: font.slice() };
      handle.registerFont(later);
      for (const face of [opened, later]) { face.bytes.fill(0); face.family = 'Changed'; }
      const png = await handle.renderPng({ slide: 0 });
      if (native) expect(isPng(png)).toBe(true);
      else {
        expect(browser.sources.length).toBeGreaterThan(0);
        for (const source of browser.sources) expect(source).toEqual(font);
      }
    } finally { handle.dispose(); }
  });
});

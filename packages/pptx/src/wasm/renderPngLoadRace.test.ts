import { afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as fonts from '@betteroffice/fonts';
import { initWasm, openPresentation } from '../index';
import { engineKind } from './engine';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = new Uint8Array(pptx);
});

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stubCanvas() {
  class Canvas {
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
  class Face { constructor(public family: string) {} load() { return Promise.resolve(this); } }
  const held: Face[] = [];
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, FontFace: scope.FontFace, fonts: scope.fonts, Image: scope.Image, createImageBitmap: scope.createImageBitmap };
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
}

describe.skipIf(engineKind === 'native')('canvas renderPng with a load that fails while another request still resolves', () => {
  test('rejects with font-unavailable and leaves no unhandled rejection', async () => {
    stubCanvas();
    const installed = fonts.bundledFaceInstalled;
    let held = 0;
    let release!: () => void;
    const gate = new Promise<void>((done) => { release = done; });
    const checking = spyOn(fonts, 'bundledFaceInstalled').mockImplementation((...args: Parameters<typeof installed>) => {
      if (!args[0].script?.startsWith('cjk-')) return installed(...args);
      held += 1;
      return gate.then(() => installed(...args));
    });
    const failing = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(() => Promise.reject(new Error('offline')));
    const stray: unknown[] = [];
    const onStray = (error: unknown) => { stray.push(error); };
    process.on('unhandledRejection', onStray);
    const handle = openPresentation(fixture, { clientId: 7304 });
    try {
      const slideId = handle.snapshot().slides[0]!.id;
      handle.addTextBox(slideId, { name: 'Han', rect: { x: 100000, y: 100000, width: 3000000, height: 800000 }, text: '\u6f22\u5b57', style: { fontFamily: 'Arial', fontSizePt: 18 } });
      const exported = handle.renderPng({ slide: 0 });
      exported.catch(() => {});
      await new Promise((settled) => setTimeout(settled, 50));
      expect(held).toBeGreaterThan(0);
      release();
      await expect(exported).rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      await new Promise((settled) => setTimeout(settled, 50));
      expect(stray).toEqual([]);
    } finally {
      process.off('unhandledRejection', onStray);
      release();
      checking.mockRestore();
      failing.mockRestore();
      handle.dispose();
    }
  });
});

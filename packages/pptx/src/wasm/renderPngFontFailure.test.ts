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

describe.skipIf(engineKind === 'native')('canvas renderPng with a bundled face that cannot load', () => {
  test('rejects with font-unavailable naming the face, and exports once the face loads', async () => {
    stubCanvas();
    const failing = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(() => Promise.reject(new Error('offline')));
    const handle = openPresentation(fixture, { clientId: 7301 });
    try {
      await expect(handle.renderPng({ slide: 0 })).rejects.toMatchObject({
        name: 'PngExportError', code: 'font-unavailable', message: expect.stringContaining(' could not load: offline'),
      });
      failing.mockRestore();
      expect(new TextDecoder().decode(await handle.renderPng({ slide: 0 }))).toBe('png');
    } finally {
      failing.mockRestore();
      handle.dispose();
    }
  });
});

describe.skipIf(engineKind === 'native')('canvas renderPng while edits keep asking for new faces', () => {
  test('rejects with font-unavailable after a bounded number of loading rounds', async () => {
    stubCanvas();
    const original = fonts.bundledFaceInstalled;
    const handle = openPresentation(fixture, { clientId: 7302 });
    const slideId = handle.snapshot().slides[0]!.id;
    let edits = 0;
    const loading = spyOn(fonts, 'bundledFaceInstalled').mockImplementation((...args: Parameters<typeof original>) => {
      edits += 1;
      handle.addTextBox(slideId, { name: `Edit ${edits}`, rect: { x: 100000, y: 100000, width: 3000000, height: 800000 }, text: `Edit ${edits}`, style: { fontFamily: `Edited Family ${edits}`, fontSizePt: 18 } });
      return original(...args);
    });
    const extra = { family: 'Supplied Only Here', bytes: new Uint8Array(await readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))) };
    const stray: unknown[] = [];
    const onStray = (error: unknown) => { stray.push(error); };
    process.on('unhandledRejection', onStray);
    try {
      await expect(handle.renderPng({ slide: 0 })).rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      await expect(handle.renderPng({ slide: 0 }, { fonts: [extra] })).rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      await new Promise((settled) => setTimeout(settled, 50));
      expect(stray).toEqual([]);
      loading.mockRestore();
      expect(new TextDecoder().decode(await handle.renderPng({ slide: 0 }))).toBe('png');
    } finally {
      process.off('unhandledRejection', onStray);
      loading.mockRestore();
      handle.dispose();
    }
  });
});

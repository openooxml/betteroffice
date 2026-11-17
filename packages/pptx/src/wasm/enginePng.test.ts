import { afterEach, describe, expect, test } from 'bun:test';
import type { SlideDisplayList, SlidePrimitive } from '../types';
import { engineKind, renderPng } from './enginePng';

const picture = {
  kind: 'image', objectId: 1, shapeId: 'shape:1', name: 'Logo', x: 10, y: 10, w: 100, h: 50, assetId: 'media:logo',
} as unknown as SlidePrimitive;

const list = { contractVersion: 1, width: 320, height: 180, background: { kind: 'solid', color: '#ffffff' }, primitives: [picture] } as SlideDisplayList;
const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>');

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function workerScope() {
  const calls = { encoded: 0, decoded: 0 };
  class Canvas {
    getContext() { return new Proxy({}, { get: () => () => undefined }); }
    convertToBlob() { calls.encoded++; return Promise.resolve(new Blob(['png'])); }
  }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, Image: scope.Image, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, Image: undefined,
    createImageBitmap: () => { calls.decoded++; return Promise.reject(new Error('no decoder in this test')); },
  });
  restore = () => { Object.assign(scope, saved); };
  return calls;
}

describe.skipIf(engineKind === 'native')('canvas renderPng pictures', () => {
  test('an SVG picture where only a document can decode it rejects instead of leaving the picture out', async () => {
    const calls = workerScope();
    const engine = { displayList: () => list, media: () => svg, fonts: () => [], doc: null, renderer: null };
    await expect(renderPng(engine, { slide: 0 }, {})).rejects.toMatchObject({
      name: 'PngExportError', code: 'canvas-unavailable', message: expect.stringContaining('@betteroffice/pptx-native'),
    });
    expect(calls).toEqual({ encoded: 0, decoded: 0 });
  });
});

import { describe, expect, test } from 'bun:test';
import { decodePresentationImage } from './image';
import { paintSlide } from './canvas';

function png(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

function rasterHeaders(): [string, Uint8Array, number, number][] {
  const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0, 7, 8, 0x10, 0, 0x20, 0, 0xff, 0xd9]);
  const gif = new Uint8Array([71, 73, 70, 56, 57, 97, 0, 0x20, 0, 0x10]);
  const bmp = new Uint8Array(54);
  const bmpView = new DataView(bmp.buffer);
  bmpView.setUint16(0, 0x4d42, true);
  bmpView.setUint32(14, 40, true);
  bmpView.setInt32(18, 8192, true);
  bmpView.setInt32(22, -4096, true);
  const webp = new Uint8Array(30);
  webp.set(new TextEncoder().encode('RIFF'));
  webp.set(new TextEncoder().encode('WEBPVP8X'), 8);
  webp.set([0xff, 0x1f, 0, 0xff, 0x0f, 0], 24);
  const exif = new Uint8Array(34);
  exif.set([69, 120, 105, 102, 0, 0, 73, 73, 42, 0, 8, 0, 0, 0]);
  const exifView = new DataView(exif.buffer);
  exifView.setUint16(14, 1, true);
  exifView.setUint16(16, 0x0112, true);
  exifView.setUint16(18, 3, true);
  exifView.setUint32(20, 1, true);
  exifView.setUint16(24, 6, true);
  const rotated = new Uint8Array(jpeg.length + exif.length + 4);
  rotated.set([0xff, 0xd8, 0xff, 0xe1, 0, exif.length + 2]);
  rotated.set(exif, 6);
  rotated.set(jpeg.subarray(2), 6 + exif.length);
  return [
    ['JPEG', jpeg, 4096, 2048], ['GIF', gif, 4096, 2048],
    ['BMP', bmp, 4096, 2048], ['WebP', webp, 4096, 2048],
    ['rotated JPEG', rotated, 2048, 4096],
  ];
}

describe('presentation image decode budget', () => {
  test.each([[8192, 4096, 4096, 2048], [4096, 8192, 2048, 4096]])(
    'caps oversized pictures before decoding (%s by %s)',
    async (width, height, resizeWidth, resizeHeight) => {
      const original = globalThis.createImageBitmap;
      const calls: (ImageBitmapOptions | undefined)[] = [];
      globalThis.createImageBitmap = ((_blob: Blob, options?: ImageBitmapOptions) => {
        calls.push(options);
        return Promise.resolve({ width: resizeWidth, height: resizeHeight } as ImageBitmap);
      }) as typeof createImageBitmap;
      try {
        await decodePresentationImage(png(width, height), 'undecodable');
        expect(calls).toEqual([{ resizeWidth, resizeHeight }]);
      } finally {
        globalThis.createImageBitmap = original;
      }
    }
  );

  test.each(rasterHeaders())('caps oversized %s pictures before decoding', async (_name, bytes, resizeWidth, resizeHeight) => {
    const original = globalThis.createImageBitmap;
    const calls: (ImageBitmapOptions | undefined)[] = [];
    globalThis.createImageBitmap = ((_blob: Blob, options?: ImageBitmapOptions) => {
      calls.push(options);
      return Promise.resolve({ width: resizeWidth, height: resizeHeight } as ImageBitmap);
    }) as typeof createImageBitmap;
    try {
      await decodePresentationImage(bytes, 'undecodable');
      expect(calls).toEqual([{ resizeWidth, resizeHeight }]);
    } finally {
      globalThis.createImageBitmap = original;
    }
  });

  test('keeps small pictures at their intrinsic size', async () => {
    const original = globalThis.createImageBitmap;
    const calls: (ImageBitmapOptions | undefined)[] = [];
    globalThis.createImageBitmap = ((_blob: Blob, options?: ImageBitmapOptions) => {
      calls.push(options);
      return Promise.resolve({ width: 320, height: 180 } as ImageBitmap);
    }) as typeof createImageBitmap;
    try {
      await decodePresentationImage(png(320, 180), 'undecodable');
      expect(calls).toEqual([undefined]);
    } finally {
      globalThis.createImageBitmap = original;
    }
  });

  test('honors an optional decode cap', async () => {
    const original = globalThis.createImageBitmap;
    let resize: ImageBitmapOptions | undefined;
    globalThis.createImageBitmap = ((_blob: Blob, options?: ImageBitmapOptions) => {
      resize = options;
      return Promise.resolve({ width: 2560, height: 1280 } as ImageBitmap);
    }) as typeof createImageBitmap;
    try {
      await decodePresentationImage(png(8192, 4096), 'undecodable', { maxDimension: 2560 });
      expect(resize).toEqual({ resizeWidth: 2560, resizeHeight: 1280 });
    } finally {
      globalThis.createImageBitmap = original;
    }
  });

  test('caps SVG dimensions before element decoding', async () => {
    const original = { image: globalThis.Image, url: URL.createObjectURL };
    let source: Blob | undefined;
    URL.createObjectURL = ((blob: Blob) => { source = blob; return 'blob:capped-svg'; }) as typeof URL.createObjectURL;
    globalThis.Image = class {
      onload: (() => void) | null = null;
      set src(_value: string) { queueMicrotask(() => this.onload?.()); }
    } as unknown as typeof Image;
    try {
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="8192" height="4096"/>';
      await decodePresentationImage(new TextEncoder().encode(svg), 'undecodable');
      const text = await source!.text();
      expect(text).toContain('width="4096" height="2048"/>');
      expect(text).toContain('viewBox="0 0 8192 4096"');
      expect(text).toContain('width:4096px!important;height:2048px!important');
    } finally {
      globalThis.Image = original.image;
      URL.createObjectURL = original.url;
    }
  });

  test('rejects oversized raster pictures without bitmap resizing', async () => {
    const original = globalThis.createImageBitmap;
    (globalThis as { createImageBitmap?: unknown }).createImageBitmap = undefined;
    try {
      await expect(decodePresentationImage(png(8192, 4096), 'undecodable')).rejects.toThrow('undecodable');
    } finally {
      globalThis.createImageBitmap = original;
    }
  });

  test.each([false, true])('preserves tile size after downsampling (effects=%s)', async (effects) => {
    const original = { bitmap: globalThis.createImageBitmap, matrix: globalThis.DOMMatrix, canvas: globalThis.OffscreenCanvas };
    const scales: number[][] = [];
    globalThis.createImageBitmap = (() => Promise.resolve({ width: 4096, height: 2048 } as ImageBitmap)) as typeof createImageBitmap;
    globalThis.DOMMatrix = class {
      translateSelf() { return this; }
      scaleSelf(x: number, y: number) { scales.push([x, y]); return this; }
    } as unknown as typeof DOMMatrix;
    globalThis.OffscreenCanvas = class {
      constructor(public width: number, public height: number) {}
      getContext() {
        return {
          drawImage() {},
          getImageData: () => ({ data: new Uint8ClampedArray(4) }),
          putImageData() {},
        };
      }
    } as unknown as typeof OffscreenCanvas;
    try {
      const source = await decodePresentationImage(png(8192, 4096), 'undecodable');
      const ctx = new Proxy({} as CanvasRenderingContext2D, {
        get: (_, key) => key === 'createPattern' ? () => ({ setTransform: () => {} }) : () => {},
        set: () => true,
      });
      await paintSlide(ctx, {
        contractVersion: 1, width: 960, height: 540,
        primitives: [{ kind: 'image', objectId: 1, name: 'Tiled', assetId: 'image',
          x: 0, y: 0, w: 960, h: 540, tile: { scaleX: 0.5, scaleY: 0.25 },
          effects: effects ? [{ kind: 'grayscale' }] : undefined }],
      }, 1, 1, { resolveImage: () => source });
      expect(scales).toEqual([[1, 0.5]]);
    } finally {
      globalThis.createImageBitmap = original.bitmap;
      globalThis.DOMMatrix = original.matrix;
      globalThis.OffscreenCanvas = original.canvas;
    }
  });
});

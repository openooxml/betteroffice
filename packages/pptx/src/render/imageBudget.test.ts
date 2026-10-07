import { describe, expect, mock, test } from 'bun:test';
import { decodePresentationImage, imageDecodeScale, presentationImageBlob } from './image';
import { paintSlide } from './canvas';

class DecodeCanvas {
  readonly drawImage = mock((_source: CanvasImageSource, ..._coordinates: number[]) => {});
  constructor(public width: number, public height: number) {}
  getContext() { return { drawImage: this.drawImage }; }
}

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

  test('bounds unparseable raster headers after bitmap decoding without resize options', async () => {
    const original = { bitmap: globalThis.createImageBitmap, canvas: globalThis.OffscreenCanvas };
    const decoded = { width: 8192, height: 4096, close: mock(() => {}) };
    const bounded = { width: 4096, height: 2048, close: mock(() => {}) };
    let decodedCount = 0;
    const decode = mock((_source: ImageBitmapSource, _options?: ImageBitmapOptions) =>
      Promise.resolve((decodedCount++ === 0 ? decoded : bounded) as ImageBitmap));
    globalThis.createImageBitmap = decode as typeof createImageBitmap;
    globalThis.OffscreenCanvas = DecodeCanvas as unknown as typeof OffscreenCanvas;
    try {
      const ico = new Uint8Array([0, 0, 1, 0, 1, 0, 32, 32, 0, 0, 1, 0, 32, 0]);
      const source = await decodePresentationImage(ico, 'undecodable');
      expect(decode).toHaveBeenCalledTimes(2);
      expect(decode.mock.calls[0]).toEqual([expect.any(Blob)]);
      const canvas = decode.mock.calls[1][0] as unknown as DecodeCanvas;
      expect(decode.mock.calls[1]).toEqual([canvas as unknown as ImageBitmapSource]);
      expect([canvas.width, canvas.height]).toEqual([4096, 2048]);
      expect(canvas.drawImage).toHaveBeenCalledWith(decoded, 0, 0, 4096, 2048);
      expect(decoded.close).toHaveBeenCalledTimes(1);
      expect(source).toBe(bounded);
      expect(imageDecodeScale(source)).toEqual({ x: 2, y: 2 });
    } finally {
      globalThis.createImageBitmap = original.bitmap;
      globalThis.OffscreenCanvas = original.canvas;
    }
  });

  test('bounds bitmaps that ignore resize options and records actual returned dimensions', async () => {
    const original = { bitmap: globalThis.createImageBitmap, canvas: globalThis.OffscreenCanvas };
    const decoded = { width: 8192, height: 4096, close: mock(() => {}) };
    const bounded = { width: 4095, height: 2047, close: mock(() => {}) };
    let decodedCount = 0;
    const decode = mock((_source: ImageBitmapSource, _options?: ImageBitmapOptions) =>
      Promise.resolve((decodedCount++ === 0 ? decoded : bounded) as ImageBitmap));
    globalThis.createImageBitmap = decode as typeof createImageBitmap;
    globalThis.OffscreenCanvas = DecodeCanvas as unknown as typeof OffscreenCanvas;
    try {
      const source = await decodePresentationImage(png(8192, 4096), 'undecodable');
      expect(decode.mock.calls[0]).toEqual([expect.any(Blob), { resizeWidth: 4096, resizeHeight: 2048 }]);
      expect(decode).toHaveBeenCalledTimes(2);
      const canvas = decode.mock.calls[1][0] as unknown as DecodeCanvas;
      expect(decode.mock.calls[1]).toEqual([canvas as unknown as ImageBitmapSource]);
      expect([canvas.width, canvas.height]).toEqual([4096, 2048]);
      expect(canvas.drawImage).toHaveBeenCalledWith(decoded, 0, 0, 4096, 2048);
      expect(decoded.close).toHaveBeenCalledTimes(1);
      expect(source).toBe(bounded);
      expect(Math.max(bounded.width, bounded.height)).toBeLessThanOrEqual(4096);
      expect(imageDecodeScale(source)).toEqual({ x: 8192 / 4095, y: 4096 / 2047 });
    } finally {
      globalThis.createImageBitmap = original.bitmap;
      globalThis.OffscreenCanvas = original.canvas;
    }
  });

  test('returns the capped canvas when bitmap conversion rejects', async () => {
    const original = { bitmap: globalThis.createImageBitmap, canvas: globalThis.OffscreenCanvas };
    const decoded = { width: 8192, height: 4096, close: mock(() => {}) };
    const decode = mock((source: ImageBitmapSource, _options?: ImageBitmapOptions) =>
      source instanceof Blob ? Promise.resolve(decoded as ImageBitmap) : Promise.reject(new Error('conversion failed')));
    try {
      globalThis.createImageBitmap = decode as typeof createImageBitmap;
      globalThis.OffscreenCanvas = DecodeCanvas as unknown as typeof OffscreenCanvas;
      const source = await decodePresentationImage(png(8192, 4096), 'undecodable');
      expect(decode).toHaveBeenCalledTimes(2);
      expect(decode.mock.calls[0]).toEqual([expect.any(Blob), { resizeWidth: 4096, resizeHeight: 2048 }]);
      const canvas = decode.mock.calls[1][0] as unknown as DecodeCanvas;
      expect(source).toBe(canvas as unknown as CanvasImageSource);
      expect([canvas.width, canvas.height]).toEqual([4096, 2048]);
      expect(canvas.drawImage).toHaveBeenCalledWith(decoded, 0, 0, 4096, 2048);
      expect(decoded.close).toHaveBeenCalledTimes(1);
      expect(imageDecodeScale(source)).toEqual({ x: 2, y: 2 });
    } finally {
      globalThis.createImageBitmap = original.bitmap;
      globalThis.OffscreenCanvas = original.canvas;
    }
  });

  test('uses the document canvas when the OffscreenCanvas constructor throws', async () => {
    const original = {
      bitmap: globalThis.createImageBitmap, canvas: globalThis.OffscreenCanvas, document: globalThis.document,
    };
    const decoded = { width: 8192, height: 4096, close: mock(() => {}) };
    const canvas = new DecodeCanvas(0, 0);
    const createElement = mock((_tag: string) => canvas);
    const decode = mock((source: ImageBitmapSource, _options?: ImageBitmapOptions) =>
      Promise.resolve(source instanceof Blob ? decoded as ImageBitmap : source as ImageBitmap));
    try {
      globalThis.createImageBitmap = decode as typeof createImageBitmap;
      globalThis.OffscreenCanvas = class {
        constructor() { throw new Error('canvas unavailable'); }
      } as unknown as typeof OffscreenCanvas;
      globalThis.document = { createElement } as unknown as Document;
      const source = await decodePresentationImage(png(8192, 4096), 'undecodable');
      expect(createElement).toHaveBeenCalledWith('canvas');
      expect(decode).toHaveBeenCalledTimes(2);
      expect(decode.mock.calls[1]).toEqual([canvas as unknown as ImageBitmapSource]);
      expect(source).toBe(canvas as unknown as CanvasImageSource);
      expect([canvas.width, canvas.height]).toEqual([4096, 2048]);
      expect(canvas.drawImage).toHaveBeenCalledWith(decoded, 0, 0, 4096, 2048);
      expect(decoded.close).toHaveBeenCalledTimes(1);
      expect(imageDecodeScale(source)).toEqual({ x: 2, y: 2 });
    } finally {
      globalThis.createImageBitmap = original.bitmap;
      globalThis.OffscreenCanvas = original.canvas;
      globalThis.document = original.document;
    }
  });

  test('uses the document canvas when the OffscreenCanvas 2D context is null', async () => {
    const original = {
      bitmap: globalThis.createImageBitmap, canvas: globalThis.OffscreenCanvas, document: globalThis.document,
    };
    const decoded = { width: 8192, height: 4096, close: mock(() => {}) };
    const canvas = new DecodeCanvas(0, 0);
    const createElement = mock((_tag: string) => canvas);
    const getContext = mock((_context: string) => null);
    const decode = mock((source: ImageBitmapSource, _options?: ImageBitmapOptions) =>
      Promise.resolve(source instanceof Blob ? decoded as ImageBitmap : source as ImageBitmap));
    try {
      globalThis.createImageBitmap = decode as typeof createImageBitmap;
      globalThis.OffscreenCanvas = class {
        constructor(public width: number, public height: number) {}
        getContext = getContext;
      } as unknown as typeof OffscreenCanvas;
      globalThis.document = { createElement } as unknown as Document;
      const source = await decodePresentationImage(png(8192, 4096), 'undecodable');
      expect(getContext).toHaveBeenCalledWith('2d');
      expect(createElement).toHaveBeenCalledWith('canvas');
      expect(decode).toHaveBeenCalledTimes(2);
      expect(decode.mock.calls[1]).toEqual([canvas as unknown as ImageBitmapSource]);
      expect(source).toBe(canvas as unknown as CanvasImageSource);
      expect([canvas.width, canvas.height]).toEqual([4096, 2048]);
      expect(canvas.drawImage).toHaveBeenCalledWith(decoded, 0, 0, 4096, 2048);
      expect(decoded.close).toHaveBeenCalledTimes(1);
      expect(imageDecodeScale(source)).toEqual({ x: 2, y: 2 });
    } finally {
      globalThis.createImageBitmap = original.bitmap;
      globalThis.OffscreenCanvas = original.canvas;
      globalThis.document = original.document;
    }
  });

  test.each(['svg', 'svg:svg'])('preserves %s element decode bytes and bounds oversized SVGs', async (root) => {
    const original = {
      image: globalThis.Image, url: URL.createObjectURL, revoke: URL.revokeObjectURL,
      bitmap: globalThis.createImageBitmap, canvas: globalThis.OffscreenCanvas,
    };
    let source: Blob | undefined;
    let element: CanvasImageSource | undefined;
    const decode = mock((canvas: ImageBitmapSource) => Promise.resolve({
      width: (canvas as OffscreenCanvas).width, height: (canvas as OffscreenCanvas).height,
    } as ImageBitmap));
    globalThis.createImageBitmap = decode as typeof createImageBitmap;
    globalThis.OffscreenCanvas = DecodeCanvas as unknown as typeof OffscreenCanvas;
    URL.createObjectURL = ((blob: Blob) => { source = blob; return 'blob:capped-svg'; }) as typeof URL.createObjectURL;
    URL.revokeObjectURL = mock(() => {});
    globalThis.Image = class {
      naturalWidth = 8192;
      naturalHeight = 4096;
      onload: (() => void) | null = null;
      set src(_value: string) {
        element = this as unknown as HTMLImageElement;
        queueMicrotask(() => this.onload?.());
      }
    } as unknown as typeof Image;
    try {
      const svg = `<${root} xmlns="http://www.w3.org/2000/svg" xmlns:svg="http://www.w3.org/2000/svg" data-note=' width="8192"' style="width:8192px!important;width:16px;height:4096px!important;height:8px"/>`;
      const bytes = new TextEncoder().encode(svg);
      const expected = presentationImageBlob(bytes);
      const bounded = await decodePresentationImage(bytes, 'undecodable');
      expect(source!.type).toBe(expected.type);
      expect(new Uint8Array(await source!.arrayBuffer())).toEqual(new Uint8Array(await expected.arrayBuffer()));
      expect(decode).toHaveBeenCalledTimes(1);
      const canvas = decode.mock.calls[0][0] as unknown as DecodeCanvas;
      expect([canvas.width, canvas.height]).toEqual([4096, 2048]);
      expect(canvas.drawImage).toHaveBeenCalledWith(element, 0, 0, 4096, 2048);
      expect([(bounded as ImageBitmap).width, (bounded as ImageBitmap).height]).toEqual([4096, 2048]);
      expect(imageDecodeScale(bounded)).toEqual({ x: 2, y: 2 });
    } finally {
      globalThis.Image = original.image;
      URL.createObjectURL = original.url;
      URL.revokeObjectURL = original.revoke;
      globalThis.createImageBitmap = original.bitmap;
      globalThis.OffscreenCanvas = original.canvas;
    }
  });

  test.each(['offscreen', 'document'])('bounds element-decoded raster pictures with a %s canvas', async (kind) => {
    const original = {
      image: globalThis.Image, url: URL.createObjectURL, revoke: URL.revokeObjectURL,
      bitmap: globalThis.createImageBitmap, canvas: globalThis.OffscreenCanvas, document: globalThis.document,
    };
    let element: CanvasImageSource | undefined;
    const canvas = new DecodeCanvas(0, 0);
    (globalThis as { createImageBitmap?: unknown }).createImageBitmap = undefined;
    if (kind === 'offscreen') globalThis.OffscreenCanvas = DecodeCanvas as unknown as typeof OffscreenCanvas;
    else {
      (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = undefined;
      globalThis.document = { createElement: mock(() => canvas) } as unknown as Document;
    }
    URL.createObjectURL = (() => 'blob:capped-raster') as typeof URL.createObjectURL;
    URL.revokeObjectURL = mock(() => {});
    globalThis.Image = class {
      naturalWidth = 8192;
      naturalHeight = 4096;
      width = 320;
      height = 180;
      onload: (() => void) | null = null;
      set src(_value: string) {
        element = this as unknown as HTMLImageElement;
        queueMicrotask(() => this.onload?.());
      }
    } as unknown as typeof Image;
    try {
      const source = await decodePresentationImage(png(8192, 4096), 'undecodable');
      const bounded = source as unknown as DecodeCanvas;
      expect(bounded).toBeInstanceOf(DecodeCanvas);
      expect([bounded.width, bounded.height]).toEqual([4096, 2048]);
      expect(bounded.drawImage).toHaveBeenCalledWith(element, 0, 0, 4096, 2048);
      expect(imageDecodeScale(source)).toEqual({ x: 2, y: 2 });
      if (kind === 'document') expect(source).toBe(canvas as unknown as CanvasImageSource);
    } finally {
      globalThis.Image = original.image;
      URL.createObjectURL = original.url;
      URL.revokeObjectURL = original.revoke;
      globalThis.createImageBitmap = original.bitmap;
      globalThis.OffscreenCanvas = original.canvas;
      globalThis.document = original.document;
    }
  });

  test.each([2, 1])('paints non-empty tile crops retaining %s original pixels', async (kept) => {
    const original = { bitmap: globalThis.createImageBitmap, matrix: globalThis.DOMMatrix, canvas: globalThis.OffscreenCanvas };
    const tiles: DecodeCanvas[] = [];
    const scales: number[][] = [];
    globalThis.createImageBitmap = (() => Promise.resolve({ width: 4096, height: 2048 } as ImageBitmap)) as typeof createImageBitmap;
    globalThis.OffscreenCanvas = class extends DecodeCanvas {
      constructor(width: number, height: number) { super(width, height); tiles.push(this); }
    } as unknown as typeof OffscreenCanvas;
    globalThis.DOMMatrix = class {
      translateSelf() { return this; }
      scaleSelf(x: number, y: number) { scales.push([x, y]); return this; }
    } as unknown as typeof DOMMatrix;
    const pattern = mock(() => ({ setTransform() {} }));
    const fill = mock(() => {});
    const ctx = new Proxy({} as CanvasRenderingContext2D, {
      get: (_, key) => key === 'createPattern' ? pattern : key === 'fillRect' ? fill : () => {},
      set: () => true,
    });
    try {
      const source = await decodePresentationImage(png(8192, 4096), 'undecodable');
      await paintSlide(ctx, {
        contractVersion: 1, width: 960, height: 540,
        primitives: [{ kind: 'image', objectId: 1, name: 'Cropped tile', assetId: 'image',
          x: 0, y: 0, w: 960, h: 540, tile: { scaleX: 1, scaleY: 1 },
          crop: { left: 4095 / 8192, right: (4097 - kept) / 8192 } }],
      }, 1, 1, { resolveImage: () => source });
      expect(tiles).toHaveLength(1);
      expect([tiles[0].width, tiles[0].height]).toEqual([1, 2048]);
      expect(tiles[0].drawImage).toHaveBeenCalledWith(source, 2047.5, 0, kept / 2, 2048, 0, 0, 1, 2048);
      expect(pattern).toHaveBeenCalledTimes(1);
      expect(fill).toHaveBeenCalledTimes(1);
      expect(scales).toEqual([[kept, 2]]);
    } finally {
      globalThis.createImageBitmap = original.bitmap;
      globalThis.DOMMatrix = original.matrix;
      globalThis.OffscreenCanvas = original.canvas;
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

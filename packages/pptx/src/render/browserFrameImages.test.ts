import { afterEach, expect, test } from 'bun:test';
import { browserFrameImages } from './browserFrameImages';

const original = Object.getOwnPropertyDescriptor(globalThis, 'createImageBitmap');
afterEach(() => {
  if (original) Object.defineProperty(globalThis, 'createImageBitmap', original);
  else Reflect.deleteProperty(globalThis, 'createImageBitmap');
});

function png(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width); view.setUint32(20, height);
  bytes[24] = 8; bytes[25] = 6;
  return bytes;
}

test('a decode limit resizes large pictures while the default cache keeps them whole', async () => {
  const calls: (ImageBitmapOptions | null)[] = [];
  Object.defineProperty(globalThis, 'createImageBitmap', { configurable: true, value: async (_source: unknown, options?: ImageBitmapOptions) => {
    calls.push(options ?? null);
    return { width: options?.resizeWidth ?? 2000, height: options?.resizeHeight ?? 1000, close() {} };
  } });
  const media = new Map([['picture', png(2000, 1000)]]);
  const thumbnails = browserFrameImages({ maxDimension: 512 }), slides = browserFrameImages();
  try {
    await thumbnails.resolve({ media })('picture');
    await slides.resolve({ media })('picture');
    expect(calls).toEqual([{ resizeWidth: 512, resizeHeight: 256 }, null]);
  } finally { thumbnails.dispose(); slides.dispose(); }
});

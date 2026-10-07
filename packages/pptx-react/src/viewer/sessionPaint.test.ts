import { expect, mock, test } from 'bun:test';
import type { PresentationFrame } from '@betteroffice/pptx';
import { paintSlide } from '@betteroffice/pptx';
import { frameImages } from './sessionPaint';

function frame(): PresentationFrame {
  const bytes = new Uint8Array(24);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, 4096);
  view.setUint32(20, 4096);
  return {
    slideIndex: 0, version: 'v1', epoch: 1,
    displayList: { contractVersion: 1, width: 960, height: 540, primitives: [] },
    media: new Map(['a', 'b', 'c'].map((id) => [id, bytes] as const)),
  };
}

test('does not call mediaBytes on a second paint of a cached picture', async () => {
  const original = globalThis.createImageBitmap;
  const bitmap = { width: 4096, height: 4096, close: mock(() => {}) };
  globalThis.createImageBitmap = (() => Promise.resolve(bitmap as ImageBitmap)) as typeof createImageBitmap;
  const images = frameImages();
  const image = frame();
  const mediaBytes = mock((id: string) => image.media.get(id)?.slice());
  const resolveImage = images.resolver(mediaBytes, 'undecodable');
  const painted = mock(() => {});
  const ctx = new Proxy({} as CanvasRenderingContext2D, {
    get: (_, key) => key === 'drawImage' ? painted : () => {},
    set: () => true,
  });
  try {
    const display = {
      ...image.displayList,
      primitives: [{ kind: 'image' as const, objectId: 1, name: 'Picture', assetId: 'a',
        x: 0, y: 0, w: 960, h: 540 }],
    };
    await paintSlide(ctx, display, 1, 1, { resolveImage });
    expect(mediaBytes).toHaveBeenCalledTimes(1);
    await paintSlide(ctx, display, 1, 1, { resolveImage });
    expect(mediaBytes).toHaveBeenCalledTimes(1);
    expect(painted).toHaveBeenCalledTimes(2);
  } finally {
    images.dispose();
    globalThis.createImageBitmap = original;
  }
});

test('evicts decoded bytes in least recently used order', async () => {
  const original = globalThis.createImageBitmap;
  const bitmaps = Array.from({ length: 4 }, () => ({ width: 4096, height: 4096, close: mock(() => {}) }));
  let decoded = 0;
  globalThis.createImageBitmap = (() => Promise.resolve(bitmaps[decoded++] as ImageBitmap)) as typeof createImageBitmap;
  const images = frameImages();
  try {
    const image = frame();
    const paint = async (id: string) => {
      const resolve = images.resolve(image);
      try { return await resolve(id); } finally { resolve.release(); }
    };
    expect(await paint('a')).toBe(bitmaps[0]);
    expect(await paint('b')).toBe(bitmaps[1]);
    expect(await paint('a')).toBe(bitmaps[0]);
    expect(await paint('c')).toBe(bitmaps[2]);
    expect(bitmaps[1].close).toHaveBeenCalledTimes(1);
    expect(bitmaps[0].close).not.toHaveBeenCalled();
    expect(await paint('b')).toBe(bitmaps[3]);
    expect(bitmaps[0].close).toHaveBeenCalledTimes(1);
  } finally {
    images.dispose();
    globalThis.createImageBitmap = original;
  }
});

test('holds a bitmap evicted by decoded bytes until every paint releases it', async () => {
  const original = globalThis.createImageBitmap;
  const bitmaps = Array.from({ length: 3 }, () => ({ width: 4096, height: 4096, close: mock(() => {}) }));
  let decoded = 0;
  globalThis.createImageBitmap = (() => Promise.resolve(bitmaps[decoded++] as ImageBitmap)) as typeof createImageBitmap;
  const images = frameImages();
  const image = frame();
  const first = images.resolve(image);
  const second = images.resolve(image);
  try {
    await first('a');
    await second('a');
    for (const id of ['b', 'c']) {
      const paint = images.resolve(image);
      await paint(id);
      paint.release();
    }
    expect(bitmaps[0].close).not.toHaveBeenCalled();
    first.release();
    expect(bitmaps[0].close).not.toHaveBeenCalled();
    second.release();
    expect(bitmaps[0].close).toHaveBeenCalledTimes(1);
  } finally {
    first.release();
    second.release();
    images.dispose();
    globalThis.createImageBitmap = original;
  }
});

test('releases local paint images after a paint failure', async () => {
  const original = globalThis.createImageBitmap;
  const bitmaps = Array.from({ length: 2 }, () => ({ width: 4096, height: 4096, close: mock(() => {}) }));
  let decoded = 0;
  globalThis.createImageBitmap = (() => Promise.resolve(bitmaps[decoded++] as ImageBitmap)) as typeof createImageBitmap;
  const images = frameImages({ maxDecodedBytes: 0 });
  try {
    const image = frame();
    const resolveImage = images.resolver((id) => image.media.get(id), 'undecodable');
    const ctx = new Proxy({} as CanvasRenderingContext2D, {
      get: (_, key) => key === 'drawImage' ? () => { throw new Error('paint failed'); } : () => {},
      set: () => true,
    });
    const display = {
      ...image.displayList,
      primitives: [{ kind: 'image' as const, objectId: 1, name: 'Picture', assetId: 'a',
        x: 0, y: 0, w: 960, h: 540 }],
    };
    for (let index = 0; index < 2; index += 1)
      await expect(paintSlide(ctx, display, 1, 1, { resolveImage })).rejects.toThrow('paint failed');
    for (const bitmap of bitmaps) expect(bitmap.close).toHaveBeenCalledTimes(1);
  } finally {
    images.dispose();
    globalThis.createImageBitmap = original;
  }
});

test('protects local paint images until painting finishes', async () => {
  const original = globalThis.createImageBitmap;
  const bitmaps = Array.from({ length: 2 }, () => ({ width: 4096, height: 4096, close: mock(() => {}) }));
  let decoded = 0;
  globalThis.createImageBitmap = (() => Promise.resolve(bitmaps[decoded++] as ImageBitmap)) as typeof createImageBitmap;
  const images = frameImages({ maxDecodedBytes: 64 * 1024 * 1024 });
  try {
    const image = frame();
    const painted: unknown[] = [];
    const ctx = new Proxy({} as CanvasRenderingContext2D, {
      get: (_, key) => key === 'drawImage' ? (source: unknown) => {
        expect(bitmaps[0].close).not.toHaveBeenCalled();
        expect(bitmaps[1].close).not.toHaveBeenCalled();
        painted.push(source);
      } : () => {},
      set: () => true,
    });
    await paintSlide(ctx, {
      ...image.displayList,
      primitives: ['a', 'b'].map((assetId, objectId) => ({ kind: 'image' as const, objectId, assetId,
        name: 'Picture', x: 0, y: 0, w: 960, h: 540 })),
    }, 1, 1, { resolveImage: images.resolver((id) => image.media.get(id), 'undecodable') });
    expect(painted).toEqual(bitmaps);
    expect(bitmaps[0].close).toHaveBeenCalledTimes(1);
    expect(bitmaps[1].close).not.toHaveBeenCalled();
  } finally {
    images.dispose();
    globalThis.createImageBitmap = original;
  }
});

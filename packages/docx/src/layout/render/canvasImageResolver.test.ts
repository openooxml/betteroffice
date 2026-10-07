import { afterEach, beforeEach, describe, expect, it } from 'bun:test';

import { createCanvasImageResolver } from './canvasImageResolver';

class FakeImage {
  static loaded: string[] = [];
  naturalWidth = 1;
  naturalHeight = 1;
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  set src(url: string) {
    FakeImage.loaded.push(url);
    queueMicrotask(() => (url.startsWith('broken') ? this.onerror?.() : this.onload?.()));
  }
}

class FakeBitmap {
  width = 2;
  height = 2;
  closed = false;
  close() {
    this.closed = true;
  }
}

describe('createCanvasImageResolver', () => {
  const originalImage = globalThis.Image;
  const originalCreateImageBitmap = globalThis.createImageBitmap;
  const originalFetch = globalThis.fetch;
  const originalCreate = URL.createObjectURL;
  const originalRevoke = URL.revokeObjectURL;
  let blobs: Blob[];
  let revoked: string[];
  let decoded: Blob[];
  let bitmaps: FakeBitmap[];

  beforeEach(() => {
    FakeImage.loaded = [];
    blobs = [];
    revoked = [];
    decoded = [];
    bitmaps = [];
    globalThis.Image = FakeImage as unknown as typeof Image;
    globalThis.createImageBitmap = (async (blob: Blob) => {
      decoded.push(blob);
      const bitmap = new FakeBitmap();
      bitmaps.push(bitmap);
      return bitmap;
    }) as unknown as typeof createImageBitmap;
    URL.createObjectURL = (blob: Blob) => {
      blobs.push(blob);
      return `blob:media-${blobs.length}`;
    };
    URL.revokeObjectURL = (url: string) => {
      revoked.push(url);
    };
  });

  afterEach(() => {
    globalThis.Image = originalImage;
    globalThis.createImageBitmap = originalCreateImageBitmap;
    globalThis.fetch = originalFetch;
    URL.createObjectURL = originalCreate;
    URL.revokeObjectURL = originalRevoke;
  });

  it('decodes once in workers', async () => {
    globalThis.Image = undefined as unknown as typeof Image;
    const resolve = createCanvasImageResolver({
      media: () => ({ bytes: new Uint8Array([1, 2, 3]), mimeType: 'image/png' }),
    });
    const first = resolve('media:0');
    expect(resolve('media:0')).toBe(first);
    expect(await first).toBe(bitmaps[0]);
    expect(await resolve('media:0')).toBe(bitmaps[0]);
    expect(bitmaps[0]!.closed).toBe(false);
    expect(decoded.map((blob) => [blob.type, blob.size])).toEqual([['image/png', 3]]);
    expect(new Uint8Array(await decoded[0]!.arrayBuffer())).toEqual(new Uint8Array([1, 2, 3]));
    expect(blobs).toEqual([]);
  });

  it('evicts and closes LRU bitmaps', async () => {
    globalThis.Image = undefined as unknown as typeof Image;
    const resolve = createCanvasImageResolver({
      media: () => ({ bytes: new Uint8Array([1]), mimeType: 'image/png' }),
      maxCacheBytes: 32,
    });
    const first = await resolve('media:0');
    const second = await resolve('media:1');
    expect(await resolve('media:0')).toBe(first);
    await resolve('media:2');
    expect(bitmaps.map((bitmap) => bitmap.closed)).toEqual([false, true, false]);
    expect(await resolve('media:0')).toBe(first);
    expect(await resolve('media:1')).not.toBe(second);
    expect(bitmaps.map((bitmap) => bitmap.closed)).toEqual([false, true, true, false]);
  });

  it('decodes local URLs in workers', async () => {
    globalThis.Image = undefined as unknown as typeof Image;
    const fetched: string[] = [];
    globalThis.fetch = (async (url: string) => {
      fetched.push(url);
      return new Response(new Blob([new Uint8Array([1])], { type: 'image/png' }));
    }) as unknown as typeof fetch;
    const resolve = createCanvasImageResolver();
    expect(await resolve('blob:embedded')).toBe(bitmaps[0]);
    expect(await resolve('data:image/png;base64,AQ==')).toBe(bitmaps[1]);
    expect(resolve('https://example.com/a.png')).toBeNull();
    expect(fetched).toEqual(['blob:embedded', 'data:image/png;base64,AQ==']);
    expect(decoded.map((blob) => [blob.type, blob.size])).toEqual([
      ['image/png', 1],
      ['image/png', 1],
    ]);
  });

  it('closes obsolete scope bitmaps', async () => {
    globalThis.Image = undefined as unknown as typeof Image;
    const complete: ((bitmap: ImageBitmap) => void)[] = [];
    globalThis.createImageBitmap = (() =>
      new Promise<ImageBitmap>((resolve) => {
        complete.push(resolve);
      })) as typeof createImageBitmap;
    let scope = 1;
    const resolve = createCanvasImageResolver({
      media: () => ({ bytes: new Uint8Array([scope]), mimeType: 'image/png' }),
      mediaScope: () => scope,
      maxCacheBytes: 16,
    });
    const firstBitmap = new FakeBitmap();
    const first = resolve('media:0');
    complete[0]!(firstBitmap as unknown as ImageBitmap);
    expect(await first).toBe(firstBitmap);
    const staleBitmap = new FakeBitmap();
    const stale = resolve('media:1');
    scope = 2;
    const currentBitmap = new FakeBitmap();
    const current = resolve('media:1');
    expect(firstBitmap.closed).toBe(true);
    complete[1]!(staleBitmap as unknown as ImageBitmap);
    expect(await stale).toBeNull();
    expect(staleBitmap.closed).toBe(true);
    expect(resolve('media:1')).toBe(current);
    complete[2]!(currentBitmap as unknown as ImageBitmap);
    expect(await current).toBe(currentBitmap);
    expect(await resolve('media:1')).toBe(currentBitmap);
    expect(currentBitmap.closed).toBe(false);
  });

  it('keeps the newest oversized bitmap cached and drawable', async () => {
    globalThis.Image = undefined as unknown as typeof Image;
    const resolve = createCanvasImageResolver({
      media: () => ({ bytes: new Uint8Array([1]), mimeType: 'image/png' }),
      maxCacheBytes: 2,
    });
    const first = await resolve('media:0');
    expect(await resolve('media:0')).toBe(first);
    expect(bitmaps).toHaveLength(1);
    await resolve('media:1');
    await Promise.resolve();
    expect(bitmaps.map((bitmap) => bitmap.closed)).toEqual([true, false]);
  });

  it('keeps concurrent bitmaps drawable', async () => {
    globalThis.Image = undefined as unknown as typeof Image;
    const resolve = createCanvasImageResolver({
      media: () => ({ bytes: new Uint8Array([1]), mimeType: 'image/png' }),
      maxCacheBytes: 16,
    });
    const first = resolve('media:0');
    const second = resolve('media:1');
    const drawFirst = Promise.resolve(first).then((image) => {
      expect(image).toBe(bitmaps[0]);
      expect(bitmaps[0]!.closed).toBe(false);
    });
    await Promise.all([drawFirst, second]);
    expect(bitmaps.map((bitmap) => bitmap.closed)).toEqual([true, false]);
  });

  it('bounds main-thread image cache', async () => {
    const resolve = createCanvasImageResolver({ maxCacheBytes: 4 });
    const first = await resolve('data:image/png;base64,AQ==');
    expect(first).toBeInstanceOf(FakeImage);
    await resolve('data:image/png;base64,Ag==');
    expect(await resolve('data:image/png;base64,AQ==')).not.toBe(first);
    expect(FakeImage.loaded).toHaveLength(3);
    expect(decoded).toEqual([]);
  });

  it('decodes a media token from its bytes once and releases the object URL', async () => {
    const reads: string[] = [];
    const resolve = createCanvasImageResolver({
      media: (token) => {
        reads.push(token);
        return token === 'media:3'
          ? { bytes: new Uint8Array([1, 2, 3]), mimeType: 'image/png' }
          : null;
      },
    });
    const first = await resolve('media:3');
    const second = await resolve('media:3');
    expect(first).toBeInstanceOf(FakeImage);
    expect(second).toBe(first);
    expect(reads).toEqual(['media:3']);
    expect(blobs.map((blob) => [blob.type, blob.size])).toEqual([['image/png', 3]]);
    expect(FakeImage.loaded).toEqual(['blob:media-1']);
    expect(revoked).toEqual(['blob:media-1']);
    expect(await resolve('media:4')).toBeNull();
    expect(reads).toEqual(['media:3', 'media:4']);
  });

  it('reads a token again once the media scope changes', async () => {
    let scope = 1;
    const reads: number[] = [];
    const resolve = createCanvasImageResolver({
      media: () => {
        reads.push(scope);
        return { bytes: new Uint8Array([scope]), mimeType: 'image/png' };
      },
      mediaScope: () => scope,
    });
    const first = await resolve('media:0');
    expect(await resolve('media:0')).toBe(first);
    scope = 2;
    expect(await resolve('media:0')).not.toBe(first);
    expect(reads).toEqual([1, 2]);
  });

  it('refuses tokens without a media source and other schemes', () => {
    expect(createCanvasImageResolver()('media:0')).toBeNull();
    const resolve = createCanvasImageResolver({ media: () => null });
    expect(resolve('https://example.com/a.png')).toBeNull();
    expect(resolve('rId5')).toBeNull();
    expect(resolve('media:01')).toBeNull();
    expect(resolve('media:x')).toBeNull();
  });

  it('paints a token whose bytes are missing or fail to decode as a grey box', async () => {
    const originalCanvas = globalThis.OffscreenCanvas;
    const fills: string[] = [];
    class FakeCanvas {
      getContext() {
        return {
          set fillStyle(value: string) {
            fills.push(value);
          },
          fillRect() {},
        };
      }
    }
    globalThis.OffscreenCanvas = FakeCanvas as unknown as typeof OffscreenCanvas;
    URL.createObjectURL = (blob: Blob) => {
      blobs.push(blob);
      return `broken-${blobs.length}`;
    };
    try {
      const resolve = createCanvasImageResolver({
        media: (token) =>
          token === 'media:1' ? { bytes: new Uint8Array([1]), mimeType: 'image/png' } : null,
      });
      expect(await resolve('media:1')).toBeInstanceOf(FakeCanvas);
      expect(revoked).toEqual(['broken-1']);
      expect(await resolve('media:2')).toBeInstanceOf(FakeCanvas);
      expect(fills).toEqual(['#e6e6e6', '#e6e6e6']);
    } finally {
      globalThis.OffscreenCanvas = originalCanvas;
    }
  });

  it('keeps decoding data URLs directly', async () => {
    const resolve = createCanvasImageResolver({ media: () => null });
    expect(await resolve('data:image/png;base64,AA==')).toBeInstanceOf(FakeImage);
    expect(FakeImage.loaded).toEqual(['data:image/png;base64,AA==']);
    expect(blobs).toEqual([]);
    expect(revoked).toEqual([]);
  });
});

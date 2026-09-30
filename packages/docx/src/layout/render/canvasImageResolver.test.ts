import { afterEach, beforeEach, describe, expect, it } from 'bun:test';

import { createCanvasImageResolver } from './canvasImageResolver';

class FakeImage {
  static loaded: string[] = [];
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  set src(url: string) {
    FakeImage.loaded.push(url);
    queueMicrotask(() => (url.startsWith('broken') ? this.onerror?.() : this.onload?.()));
  }
}

describe('createCanvasImageResolver', () => {
  const originalImage = globalThis.Image;
  const originalCreate = URL.createObjectURL;
  const originalRevoke = URL.revokeObjectURL;
  let blobs: Blob[];
  let revoked: string[];

  beforeEach(() => {
    FakeImage.loaded = [];
    blobs = [];
    revoked = [];
    globalThis.Image = FakeImage as unknown as typeof Image;
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
    URL.createObjectURL = originalCreate;
    URL.revokeObjectURL = originalRevoke;
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

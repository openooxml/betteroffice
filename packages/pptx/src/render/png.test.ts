import { afterEach, describe, expect, test } from 'bun:test';
import type { SlideDisplayList } from '../types';
import { slideToPng } from './png';

const list: SlideDisplayList = {
  contractVersion: 1,
  width: 320,
  height: 180,
  background: { kind: 'solid', color: '#ffffff' },
  primitives: [
    {
      kind: 'shape',
      objectId: 1,
      shapeId: 'shape:1',
      name: 'Card',
      x: 20,
      y: 20,
      w: 280,
      h: 140,
      geometry: 'rect',
      path: [
        { type: 'move', x: 0, y: 0 },
        { type: 'line', x: 1, y: 0 },
        { type: 'line', x: 1, y: 1 },
        { type: 'line', x: 0, y: 1 },
        { type: 'close' },
      ],
      fill: { kind: 'solid', color: '#3b82f6' },
    },
  ],
};

/** Records the size it was asked for and the blob it was asked to encode. */
function stubOffscreenCanvas(): { sizes: [number, number][] } {
  const sizes: [number, number][] = [];
  const original = globalThis.OffscreenCanvas;
  class Stub {
    constructor(width: number, height: number) {
      sizes.push([width, height]);
    }
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'createLinearGradient' || property === 'createRadialGradient') {
            return () => ({ addColorStop: () => undefined });
          }
          if (property in target) return target[property as string];
          return () => undefined;
        },
        set(target, property, value) {
          target[property as string] = value;
          return true;
        },
      });
    }
    convertToBlob(options: { type: string }) {
      return Promise.resolve(new Blob(['png'], { type: options.type }));
    }
  }
  (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = Stub;
  restore = () => {
    (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = original;
  };
  return { sizes };
}

let restore: (() => void) | undefined;

afterEach(() => {
  restore?.();
  restore = undefined;
});

describe('slideToPng', () => {
  test('encodes a png at the slide size', async () => {
    const { sizes } = stubOffscreenCanvas();
    const blob = await slideToPng(list);
    expect(sizes).toEqual([[320, 180]]);
    expect(blob.type).toBe('image/png');
  });

  test('scales the surface', async () => {
    const { sizes } = stubOffscreenCanvas();
    await slideToPng(list, { scale: 2 });
    expect(sizes).toEqual([[640, 360]]);
  });

  test('refuses a scale that is not positive', async () => {
    stubOffscreenCanvas();
    await expect(slideToPng(list, { scale: 0 })).rejects.toThrow('finite and positive');
  });
});

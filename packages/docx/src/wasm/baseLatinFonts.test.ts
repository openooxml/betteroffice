import { expect, test } from 'bun:test';
import { installBaseLatinFace } from './baseLatinFonts';

async function withFonts(run: (faces: Set<unknown>, descriptors: FontFaceDescriptors[], fail: () => void) => Promise<void>) {
  const previous = ['document', 'FontFace'].map(key => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  const faces = new Set<unknown>();
  const descriptors: FontFaceDescriptors[] = [];
  let failing = false;
  class Face {
    constructor(readonly family: string, _bytes: ArrayBuffer, descriptor: FontFaceDescriptors) { descriptors.push(descriptor); }
    async load() {
      if (failing) { failing = false; throw new Error('font unavailable'); }
      return this;
    }
  }
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts: faces } });
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
  try { await run(faces, descriptors, () => { failing = true; }); }
  finally {
    for (const [key, descriptor] of previous) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
}

test('only the requested face is installed and concurrent requests share registration', async () => {
  await withFonts(async (faces, descriptors) => {
    await Promise.all([installBaseLatinFace(new ArrayBuffer(4), true, false), installBaseLatinFace(new ArrayBuffer(4), true, false)]);
    expect(descriptors).toEqual([{ weight: '700', style: 'normal' }]);
    expect(faces.size).toBe(1);
  });
});

test('a failed browser registration does not break preparation and stays retryable', async () => {
  await withFonts(async (faces, descriptors, fail) => {
    fail();
    await expect(installBaseLatinFace(new ArrayBuffer(4), false, true)).resolves.toBeUndefined();
    expect(faces.size).toBe(0);
    await installBaseLatinFace(new ArrayBuffer(4), false, true);
    expect(faces.size).toBe(1);
    expect(descriptors).toHaveLength(2);
  });
});

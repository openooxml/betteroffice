import { expect, test } from 'bun:test';
import '../fontTraits';
import { installDefaultFonts, installFonts } from './peerlessFonts';

test('app ownership suppresses defaults both before and after browser loading', async () => {
  const descriptors = ['FontFace', 'document'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  const added: Face[] = [];
  const created: Face[] = [];
  const owned = new Set<string>();
  let finish!: () => void;
  const pending = new Promise<void>((resolve) => { finish = resolve; });
  class Face {
    constructor(readonly family: string, readonly source: ArrayBuffer) { created.push(this); }
    async load() { if (new Uint8Array(this.source)[0] === 1) await pending; return this; }
  }
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts: { add: (face: Face) => added.push(face) } } });
  try {
    const installed: FontFace[] = [];
    const allowed = (family: string) => !owned.has(family.trim().toLowerCase());
    const font = { family: 'Cambria', bold: false, italic: false, bytes: Uint8Array.of(1).buffer };
    const loading = installDefaultFonts([font], installed, () => true, () => {}, allowed);
    for (let tick = 0; tick < 20 && !created.length; tick++) await new Promise((resolve) => setTimeout(resolve, 0));
    expect(created).toHaveLength(1);
    owned.add('cambria');
    await installFonts([{ family: 'Cambria', bytes: Uint8Array.of(2) }], installed, () => true, () => {});
    finish();
    expect(await loading).toEqual([]);
    expect(added.map((face) => new Uint8Array(face.source)[0])).toEqual([2]);
    expect(await installDefaultFonts([{ ...font, family: ' CAMBRIA ', bold: true }], installed,
      () => true, () => {}, allowed)).toEqual([]);
    expect(created).toHaveLength(2);
  } finally {
    finish();
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
});

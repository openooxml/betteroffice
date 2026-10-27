import { expect, spyOn, test } from 'bun:test';
import { FONT_FALLBACK_MESSAGE } from '../fontFaces';
import { readFontTraits } from '../fontTraits';
import { installBrowserFonts, installFonts } from './peerlessFonts';
import { installFonts as installViewerFonts } from '../viewer/sessionPaint';

function sfnt({ weight = 400, selection = 0, macStyle = 0, os2 = true, signature = 0x00010000, collection = false } = {}): Uint8Array {
  const start = collection ? 20 : 0;
  const bytes = new Uint8Array(start + 164);
  const view = new DataView(bytes.buffer);
  if (collection) {
    view.setUint32(0, 0x74746366);
    view.setUint32(4, 0x00010000);
    view.setUint32(8, 1);
    view.setUint32(12, start);
  }
  view.setUint32(start, signature);
  view.setUint16(start + 4, os2 ? 2 : 1);
  view.setUint32(start + 12, 0x68656164);
  view.setUint32(start + 20, start + 44);
  view.setUint32(start + 24, 54);
  view.setUint16(start + 44 + 44, macStyle);
  if (os2) {
    view.setUint32(start + 28, 0x4f532f32);
    view.setUint32(start + 36, start + 100);
    view.setUint32(start + 40, 64);
    view.setUint16(start + 100 + 4, weight);
    view.setUint16(start + 100 + 62, selection);
  }
  return bytes;
}

test('reads selection flags while keeping the weight class authoritative', () => {
  expect(readFontTraits(sfnt({ weight: 600, selection: 32 }))).toEqual({
    weight: 600, bold: true, italic: false, oblique: false,
  });
});

test('installs fonts using the face weight and style instead of request flags', async () => {
  const descriptors = ['FontFace', 'document'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  const created: { family: string; descriptors: FontFaceDescriptors }[] = [];
  const added: FontFace[] = [];
  class Face {
    constructor(readonly family: string, _source: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {
      created.push(this);
    }
    async load() { return this; }
  }
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts: { add: (face: FontFace) => added.push(face) } } });
  try {
    const malformed = sfnt();
    new DataView(malformed.buffer).setUint32(36, malformed.length - 1);
    const truncatedDirectory = sfnt();
    new DataView(truncatedDirectory.buffer).setUint16(4, 100);
    const malformedCollection = sfnt({ collection: true });
    new DataView(malformedCollection.buffer).setUint32(12, malformedCollection.length - 1);
    const shortOs2 = sfnt();
    new DataView(shortOs2.buffer).setUint32(40, 63);
    const shortHead = sfnt({ os2: false });
    new DataView(shortHead.buffer).setUint32(24, 45);
    const missingTables = sfnt();
    new DataView(missingTables.buffer).setUint16(4, 0);
    const first = sfnt({ collection: true, weight: 700, selection: 1 });
    const second = sfnt();
    const secondView = new DataView(second.buffer);
    secondView.setUint32(20, first.length + 44);
    secondView.setUint32(36, first.length + 100);
    const collection = new Uint8Array(first.length + second.length);
    collection.set(first);
    collection.set(second, first.length);
    const collectionView = new DataView(collection.buffer);
    collectionView.setUint32(8, 2);
    collectionView.setUint32(16, first.length);
    const padded = new Uint8Array(180);
    padded.set(sfnt(), 8);
    const cases: {
      name: string;
      bytes: Uint8Array;
      bold?: boolean;
      italic?: boolean;
      weight: string;
      style: string;
    }[] = [
      { name: 'Regular requested bold', bytes: sfnt(), bold: true, weight: '400', style: 'normal' },
      { name: 'Bold', bytes: sfnt({ weight: 700, selection: 32 }), bold: true, weight: '700', style: 'normal' },
      { name: 'Italic', bytes: sfnt({ selection: 1 }), weight: '400', style: 'italic' },
      { name: 'Oblique', bytes: sfnt({ selection: 512 }), weight: '400', style: 'oblique' },
      { name: 'Italic and oblique', bytes: sfnt({ selection: 513 }), weight: '400', style: 'italic' },
      { name: 'Collection', bytes: collection, weight: '700', style: 'italic' },
      { name: 'True', bytes: sfnt({ signature: 0x74727565 }), bold: true, weight: '400', style: 'normal' },
      { name: 'OpenType', bytes: sfnt({ signature: 0x4f54544f, weight: 500 }), weight: '500', style: 'normal' },
      { name: 'Head fallback', bytes: sfnt({ os2: false, macStyle: 3 }), weight: '700', style: 'italic' },
      { name: 'Head regular', bytes: sfnt({ os2: false }), bold: true, weight: '400', style: 'normal' },
      { name: 'Head bold', bytes: sfnt({ os2: false, macStyle: 1 }), weight: '700', style: 'normal' },
      { name: 'Head italic', bytes: sfnt({ os2: false, macStyle: 2 }), weight: '400', style: 'italic' },
      { name: 'OS2 precedence', bytes: sfnt({ macStyle: 3 }), weight: '400', style: 'normal' },
      { name: 'Minimum weight', bytes: sfnt({ weight: 0 }), weight: '1', style: 'normal' },
      { name: 'Maximum weight', bytes: sfnt({ weight: 65535 }), weight: '1000', style: 'normal' },
      { name: 'Byte view', bytes: padded.subarray(8, 172), bold: true, weight: '400', style: 'normal' },
      ...[
        Uint8Array.of(1),
        Uint8Array.of(0x77, 0x4f, 0x46, 0x46),
        Uint8Array.of(0x77, 0x4f, 0x46, 0x32),
        malformed, truncatedDirectory, malformedCollection, shortOs2, shortHead, missingTables,
        sfnt({ signature: 0 }), sfnt().subarray(0, 120),
      ].map((bytes, index) => ({ name: `Unreadable ${index}`, bytes, bold: true, italic: true, weight: '700', style: 'italic' })),
    ];
    const fonts = cases.map(({ name: family, bytes, bold, italic }) => ({ family, bytes, bold, italic }));
    const installed = await installBrowserFonts(fonts);
    expect(created).toEqual(cases.map(({ name: family, weight, style }) => ({ family, descriptors: { weight, style } })));
    expect(added).toEqual(installed);
    expect(installed).toHaveLength(cases.length);
  } finally {
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
});

test('deduplicates resolved descriptors within each install call', async () => {
  const descriptors = ['FontFace', 'document'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  const added: FontFace[] = [];
  let loads = 0;
  class Face {
    constructor(readonly family: string) {}
    async load() { loads++; return this; }
  }
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts: { add: (face: FontFace) => added.push(face) } } });
  try {
    const bytes = sfnt();
    const fonts = [
      { family: 'Test', bytes },
      { family: 'Test', bytes: bytes.slice(), bold: true, italic: true },
      { family: 'Test', bytes: sfnt({ weight: 700 }) },
      { family: 'Test', bytes: sfnt({ selection: 1 }) },
      { family: 'Other', bytes },
    ];
    expect(await installBrowserFonts(fonts)).toHaveLength(4);
    expect(loads).toBe(4);
    expect(await installBrowserFonts(fonts)).toHaveLength(4);
    expect(loads).toBe(8);
    expect(added).toHaveLength(8);
  } finally {
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
});

test('a face that fails to load does not hide a later face with the same descriptors', async () => {
  const descriptors = ['FontFace', 'document'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
  const added: FontFace[] = [];
  const broken = sfnt();
  const working = broken.slice();
  working[16] ^= 1;
  class Face {
    constructor(readonly family: string, readonly source: ArrayBuffer) {}
    async load() {
      if (new Uint8Array(this.source).every((byte, index) => byte === broken[index])) throw new Error('Font blocked');
      return this;
    }
  }
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts: { add: (face: FontFace) => added.push(face) } } });
  let reports = 0;
  try {
    const installed = await installBrowserFonts([
      { family: 'Test', bytes: broken },
      { family: 'Test', bytes: working },
    ], () => true, () => { reports++; });
    expect(installed).toHaveLength(1);
    expect(added).toEqual(installed);
    expect(new Uint8Array((installed[0] as unknown as Face).source)).toEqual(working);
    expect(reports).toBe(0);
    expect(await installBrowserFonts([{ family: 'Test', bytes: broken }], () => true, () => { reports++; })).toHaveLength(0);
    expect(reports).toBe(1);
  } finally {
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  }
});

for (const mode of ['worker', 'viewer', 'in-thread'] as const) {
  test(`${mode} installs successful faces and reports rejected faces once`, async () => {
    const descriptors = ['FontFace', 'document'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
    const added: FontFace[] = [];
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    class Face {
      constructor(readonly family: string) {}
      async load() {
        if (this.family.startsWith('Failed')) throw new Error('Font blocked');
        return this;
      }
    }
    Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts: { add: (face: FontFace) => added.push(face) } } });
    try {
      const fonts = ['Regular', 'Failed', 'Failed', 'Failed too', 'Bold'].map((family) => ({ family, bytes: Uint8Array.of(1) }));
      let installed: FontFace[] = [];
      if (mode === 'in-thread') installed = await installBrowserFonts(fonts);
      else await (mode === 'worker' ? installFonts : installViewerFonts)(fonts, installed, () => true);
      expect(installed.map((face) => face.family)).toEqual(['Regular', 'Bold']);
      expect(added).toEqual(installed);
      expect(warn).toHaveBeenCalledTimes(1);
      expect(warn).toHaveBeenCalledWith(FONT_FALLBACK_MESSAGE);
    } finally {
      warn.mockRestore();
      for (const [key, descriptor] of descriptors) {
        if (descriptor) Object.defineProperty(globalThis, key, descriptor);
        else Reflect.deleteProperty(globalThis, key);
      }
    }
  });
}

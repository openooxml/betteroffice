import { expect, spyOn, test } from 'bun:test';
import { FONT_FALLBACK_MESSAGE } from '../fontFaces';
import { installBrowserFonts, installFonts } from './peerlessFonts';
import { installFonts as installViewerFonts } from '../viewer/sessionPaint';

for (const mode of ['worker', 'viewer', 'in-thread'] as const) {
  test(`${mode} installs successful faces when one face rejects`, async () => {
    const descriptors = ['FontFace', 'document'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
    const added: FontFace[] = [];
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    class Face {
      constructor(readonly family: string) {}
      async load() {
        if (this.family === 'Failed') throw new Error('Font blocked');
        return this;
      }
    }
    Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts: { add: (face: FontFace) => added.push(face) } } });
    try {
      const fonts = ['Regular', 'Failed', 'Bold'].map((family) => ({ family, bytes: Uint8Array.of(1) }));
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

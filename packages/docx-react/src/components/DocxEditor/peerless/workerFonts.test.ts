import { afterEach, expect, spyOn, test } from 'bun:test';
import * as fontManager from '@betteroffice/fonts';
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import { customWorkerFontFaces, fontEnvironmentForRegistrations, prepareWorkerFontEnvironment } from './workerFonts';

afterEach(() => { mockFetch?.mockRestore(); configureDefaultFonts({}); });
let mockFetch: ReturnType<typeof spyOn> | undefined;

test('custom WOFF and WOFF2 bytes pass unchanged to worker preparation', async () => {
  for (const signature of [0x774f4646, 0x774f4632]) {
    const data = new ArrayBuffer(16);
    new DataView(data).setUint32(0, signature);
    mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(data));
    const faces = await customWorkerFontFaces([{ family: 'Custom', src: '/face', weight: 700 }], new AbortController().signal);
    expect(faces).toMatchObject([{ family: 'Custom', weight: 'bold', data }]);
    mockFetch.mockRestore();
  }
});

test('captured registrations become face slots with style chains and substitute identities', () => {
  const environment = fontEnvironmentForRegistrations([Uint8Array.of(1), { family: 'Custom', substituteOf: 0 }],
    { 'custom|1|1': [0, 1] }, [{ key: 'custom|1|1', family: 'Custom', bold: true, italic: true }], ['Custom']);
  expect(environment).toEqual({ faces: [{ family: 'Custom', data: Uint8Array.of(1), weight: 'bold', style: 'italic' },
    { family: 'Custom', substituteOf: 0 }], chains: { 'custom|1|1': [0, 1] }, families: ['Custom'] });
});

for (const embedded of [false, true]) {
  test(`worker measurement puts URL faces before ${embedded ? 'embeds' : 'manager defaults'}`, async () => {
    const data = new Uint8Array([1, 2]).buffer;
    const custom = new Uint8Array([3, 4]).buffer;
    const bundled = new Uint8Array([5, 6]).buffer;
    const load = async () => bundled;
    const create = spyOn(fontManager, 'createFontProvider').mockReturnValue({
      resolve: () => load, resolveFamily: () => load,
      resolveScriptFallback: () => undefined, resolveLastResort: () => load,
    });
    mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(custom));
    try {
      const environment = await prepareWorkerFontEnvironment({
        fonts: [{ family: 'Custom', src: '/fonts/custom.ttf' }],
        embedded: embedded ? [{ family: 'Custom', data, weight: 'normal', style: 'normal', subsetted: false }] : [],
        requirements: [{ key: 'custom|0|0', family: 'Custom', bold: false, italic: false }],
        families: ['Custom'], provider: undefined, compat: undefined, signal: new AbortController().signal,
      });
      expect(environment.faces[0]).toMatchObject({ family: 'Custom', data: new Uint8Array(custom) });
      expect(environment.faces.filter(face => 'data' in face).map(face => face.data)).toEqual([
        new Uint8Array(custom), new Uint8Array(bundled),
      ]);
      expect(environment.chains['custom|0|0']).toEqual([0, 2, 3]);
      expect(create).toHaveBeenCalledTimes(1);
      expect(mockFetch).toHaveBeenCalledTimes(1);
    } finally { create.mockRestore(); }
  });
}

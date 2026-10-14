import { afterEach, expect, spyOn, test } from 'bun:test';
import { customWorkerFontFaces, fontEnvironmentForRegistrations } from './workerFonts';

afterEach(() => { mockFetch?.mockRestore(); });
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

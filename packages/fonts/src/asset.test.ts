import { expect, spyOn, test } from 'bun:test';
import { readFontAsset } from './asset';

test('slice files a bundler inlined as data: URLs decode without a fetch', async () => {
  const fetch = spyOn(globalThis, 'fetch');
  try {
    const binary = new Uint8Array([0, 1, 127, 128, 255, 40, 37]);
    const base64 = btoa(String.fromCharCode(...binary));
    expect(new Uint8Array(await readFontAsset('c003.zst', new URL(`data:application/octet-stream;base64,${base64}`)))).toEqual(binary);
    expect(new TextDecoder().decode(await readFontAsset('slices.json', new URL('data:application/json,%7B%22format%22%3A1%7D'))))
      .toBe('{"format":1}');
    expect(fetch).not.toHaveBeenCalled();
    expect((await readFontAsset('index.ts', new URL('./index.ts', import.meta.url))).byteLength).toBeGreaterThan(0);
  } finally {
    fetch.mockRestore();
  }
});

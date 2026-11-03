import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { BASE_LATIN_FAMILY, BASE_LATIN_FONTS, createFontProvider, loadBundledFontBytes } from './index';
import { createFontProvider as createCdnFontProvider } from './cdn';
import { fontProvider } from './provider';
import { fontAssetResponse } from './test-assets';

test('base faces use stable per-face loaders without reading bytes during resolution', async () => {
  const loaded: string[] = [];
  const provider = fontProvider(async (face) => { loaded.push(face.file); return new ArrayBuffer(0); });
  const loader = provider.resolveBaseLatin!(false, false)!;
  expect(provider.resolveBaseLatin!(false, false)).toBe(loader);
  expect(provider.resolveFamily(BASE_LATIN_FAMILY, false, false)).toBe(loader);
  expect(loaded).toEqual([]);
  await loader();
  expect(loaded).toEqual(['BetterOfficeBaseLatin-Regular.ttf.zst']);
});

test('package assets decode to the native faces without a network dependency', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (() => { throw new Error('unexpected network request'); }) as unknown as typeof fetch;
  try {
    const provider = createFontProvider();
    for (const face of BASE_LATIN_FONTS) {
      const bytes = await provider.resolveBaseLatin!(face.weight === 700, face.style === 'italic')!();
      const source = readFileSync(new URL(`../../../crates/base-latin/fonts/${face.file.slice(0, -4)}`, import.meta.url));
      expect(new Uint8Array(bytes)).toEqual(new Uint8Array(source));
    }
  } finally { globalThis.fetch = originalFetch; }
});

for (const source of ['local', 'cdn'] as const) {
  test(`${source} base assets respect an explicit asset root and use the existing decoder`, async () => {
    const originalFetch = globalThis.fetch;
    const requested: string[] = [];
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      requested.push(String(input));
      return fontAssetResponse(input);
    }) as typeof fetch;
    try {
      const baseUrl = `https://assets.example/${source}/`;
      const provider = source === 'local' ? createFontProvider({ baseUrl }) : createCdnFontProvider({ baseUrl });
      const bytes = await provider.resolveBaseLatin!(false, false)!();
      expect(bytes.byteLength).toBe(BASE_LATIN_FONTS[0]!.byteLength);
      expect(requested[0]).toBe(`${baseUrl}BetterOfficeBaseLatin-Regular.ttf.zst`);
      expect(requested.every(url => url.startsWith(baseUrl))).toBe(true);
    } finally { globalThis.fetch = originalFetch; }
  });
}

test('base assets retain manifest validation with a consumer decoder', async () => {
  const face = BASE_LATIN_FONTS[0]!;
  await expect(loadBundledFontBytes(face, undefined, {
    decoder: async () => ({ font: new ArrayBuffer(face.byteLength) }),
  })).rejects.toThrow('SHA-256 mismatch');
});

test('self-hosted base faces exist beside the other flat package assets', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const parts = new URL(String(input)).pathname.split('/');
    const file = parts[parts.length - 1]!;
    return new Response(readFileSync(new URL(`../assets/${file}`, import.meta.url)));
  }) as typeof fetch;
  try {
    const provider = createFontProvider({ baseUrl: 'https://assets.example/flat/' });
    for (const face of BASE_LATIN_FONTS) {
      const bytes = await provider.resolveBaseLatin!(face.weight === 700, face.style === 'italic')!();
      expect(bytes.byteLength).toBe(face.byteLength);
    }
    const manifest = JSON.parse(readFileSync(new URL('../assets/base-latin.json', import.meta.url), 'utf8'));
    expect(manifest.faces.map((face: { file: string }) => face.file)).toEqual(BASE_LATIN_FONTS.map(face => face.file));
  } finally { globalThis.fetch = originalFetch; }
});

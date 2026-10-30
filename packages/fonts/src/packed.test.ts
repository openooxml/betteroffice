import { afterEach, expect, mock, test } from 'bun:test';
import { createHash } from 'node:crypto';
import type { BundledFontFace } from './manifest';

const regular = new Uint8Array([0, 1, 0, 0, 10]);
const bold = new Uint8Array([0, 1, 0, 0, 20]);
const payload = new Uint8Array([91, 92, 93]);
const decode = mock(async (
  input: ArrayBuffer | Uint8Array,
  dictionary?: ArrayBuffer | Uint8Array,
  options?: { codec?: string; wasmUrl?: URL },
) => {
  const marker = new Uint8Array(input instanceof Uint8Array ? input : new Uint8Array(input))[0];
  if (marker === 1) expect(dictionary).toBeUndefined();
  else expect(new Uint8Array(dictionary as ArrayBuffer))
    .toEqual(options?.codec === 'woff2t+zstd' ? payload : regular);
  return {
    font: (marker === 1 ? regular : bold).slice().buffer,
    ...(options?.codec === 'woff2t+zstd' ? { payload: payload.slice().buffer } : {}),
  };
});
mock.module('./decode', () => ({ decodeFontFace: decode }));
const { loadPackedFontBytes, readPackManifest } = await import('./packed');

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
  decode.mockClear();
});

let sequence = 0;
function fixture(codec: 'zstd' | 'woff2t+zstd' = 'zstd') {
  const manifest = {
    format: 1, schema: 1, pack: 'synthetic', version: '1.2.3',
    faces: [regular, bold].map((bytes, index) => ({
      file: `Synthetic-${index === 0 ? 'Regular' : 'Bold'}.ttf.zst`,
      source: `Synthetic-${index === 0 ? 'Regular' : 'Bold'}.ttf`,
      family: 'Synthetic', codec, encodedLength: 1, byteLength: bytes.byteLength,
      sha256: createHash('sha256').update(bytes).digest('hex'),
      dict: index === 0 ? null : 'Synthetic-Regular.ttf.zst',
    })),
  };
  const face: BundledFontFace = {
    family: 'Synthetic', file: manifest.faces[1].file, weight: 700,
    style: 'normal', byteLength: bold.byteLength,
  };
  const requested: string[] = [];
  const base = `https://packed-${++sequence}.example/`;
  const resolveUrl = async (face: BundledFontFace) => new URL(face.file, base);
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requested.push(String(input));
    return new Response(new Uint8Array([String(input).includes('-Regular.') ? 1 : 2]));
  }) as typeof fetch;
  const load = (target = face) => loadPackedFontBytes(target, resolveUrl, undefined, undefined, manifest);
  return { manifest, face, requested, load, resolveUrl };
}

for (const codec of ['zstd', 'woff2t+zstd'] as const) {
  test(`${codec} loads its Regular prerequisite and returns transferable copies`, async () => {
    const { face, manifest, requested, load } = fixture(codec);
    const normal = { ...face, file: manifest.faces[0].file, byteLength: regular.byteLength };
    expect(requested).toEqual([]);
    const [first, second, prerequisite] = await Promise.all([load(), load(), load(normal)]);
    expect(new Uint8Array(first)).toEqual(bold);
    expect(first).not.toBe(second);
    expect(decode).toHaveBeenCalledTimes(2);
    expect(requested).toHaveLength(2);
    structuredClone(first, { transfer: [first] });
    structuredClone(prerequisite, { transfer: [prerequisite] });
    new Uint8Array(second).fill(0);
    expect(new Uint8Array(await load())).toEqual(bold);
    expect(new Uint8Array(await load(normal))).toEqual(regular);
    expect(decode).toHaveBeenCalledTimes(2);
  });
}

test('ignores unknown manifest and face fields, including slices, while loading unchanged', async () => {
  const { manifest, face, resolveUrl } = fixture();
  const extended = {
    ...manifest, future: { version: 99 },
    faces: manifest.faces.map((entry) => ({
      ...entry, future: true,
      slices: { skeleton: { file: 'unused.zst' }, chunks: [{ unknown: 123 }] },
    })),
  };
  const parsed = readPackManifest(extended);
  expect(parsed).toBe(extended);
  expect(new Uint8Array(await loadPackedFontBytes(face, resolveUrl, undefined, undefined, parsed)))
    .toEqual(bold);
  expect(extended.faces[0].slices.skeleton.file).toBe('unused.zst');
});

test('rejects a same-length hash mismatch and retries failed decodes', async () => {
  const { load } = fixture();
  decode.mockImplementationOnce(async () => ({ font: new Uint8Array(regular.byteLength).buffer }));
  await expect(load()).rejects.toThrow('SHA-256 mismatch');
  expect(new Uint8Array(await load())).toEqual(bold);
});

test('validates encoded length before decoding', async () => {
  const { manifest, load } = fixture();
  manifest.faces[1].encodedLength = 2;
  await expect(load()).rejects.toThrow('expected 2');
});

test('rejects missing or incompatible dictionary metadata', async () => {
  for (const codec of ['zstd', 'woff2t+zstd'] as const) {
    const { manifest, load } = fixture(codec);
    manifest.faces[0].codec = codec === 'zstd' ? 'woff2t+zstd' : 'zstd';
    await expect(load()).rejects.toThrow('Missing bundled font dictionary');
  }
});

test('keeps decoded dictionaries isolated by asset root', async () => {
  const first = fixture();
  const second = fixture();
  await first.load();
  await second.load();
  expect(decode).toHaveBeenCalledTimes(4);
});

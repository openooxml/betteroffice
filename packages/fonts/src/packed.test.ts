import { afterEach, expect, mock, test } from 'bun:test';
import { createHash } from 'node:crypto';
import type { BundledFontFace } from './manifest';
import type { FontDecoder } from './decode';

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
  const load = (target = face, runtime?: { decoder?: FontDecoder }) =>
    loadPackedFontBytes(target, resolveUrl, undefined, undefined, manifest, runtime);
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

for (const codec of ['zstd', 'woff2t+zstd'] as const) {
  test(`bound ${codec} decoder loads Regular and Bold without the standalone decoder`, async () => {
    const { load, requested, face, manifest } = fixture(codec);
    const engine = mock<FontDecoder>(async (input, options) => {
      expect(options.codec).toBe(codec);
      expect(options.maxOutput).toBe(64 * 1024 * 1024);
      const marker = new Uint8Array(input instanceof Uint8Array ? input : new Uint8Array(input))[0];
      if (marker === 1) expect(options.dict).toBeUndefined();
      else {
        expect(new Uint8Array(options.dict as ArrayBuffer)).toEqual(codec === 'woff2t+zstd' ? payload : regular);
        structuredClone(options.dict, { transfer: [options.dict as ArrayBuffer] });
      }
      return { font: (marker === 1 ? regular : bold).slice().buffer,
        ...(codec === 'woff2t+zstd' ? { payload: payload.slice().buffer } : {}) };
    });
    const runtime = { decoder: engine };
    const result = await load(face, runtime);
    expect(new Uint8Array(result)).toEqual(bold);
    expect(result.byteLength).toBe(bold.byteLength);
    structuredClone(result, { transfer: [result] });
    expect(new Uint8Array(await load({ ...face, file: manifest.faces[0].file }, runtime))).toEqual(regular);
    expect(engine).toHaveBeenCalledTimes(2);
    expect(decode).not.toHaveBeenCalled();
    expect(requested.filter(url => url.endsWith('fontpack.wasm'))).toEqual([]);
  });
}

test('bound decoder errors propagate without retrying the standalone decoder', async () => {
  const { load, face, requested } = fixture();
  const failure = new Error('engine decode failed');
  const decoder = mock<FontDecoder>(async () => { throw failure; });
  await expect(load(face, { decoder })).rejects.toBe(failure);
  expect(decode).not.toHaveBeenCalled();
  expect(requested.filter(url => url.endsWith('fontpack.wasm'))).toEqual([]);
});

test('unbound and distinct bound providers cannot share an in-flight decode', async () => {
  const { load, face } = fixture();
  let release!: () => void;
  const waiting = new Promise<void>(resolve => { release = resolve; });
  const first = mock<FontDecoder>(async (input) => {
    await waiting;
    return { font: (new Uint8Array(input as ArrayBuffer)[0] === 1 ? regular : bold).slice().buffer };
  });
  const second = mock<FontDecoder>(async (input) => ({
    font: (new Uint8Array(input as ArrayBuffer)[0] === 1 ? regular : bold).slice().buffer,
  }));
  const bound = load(face, { decoder: first });
  expect(new Uint8Array(await load())).toEqual(bold);
  expect(decode).toHaveBeenCalledTimes(2);
  expect(new Uint8Array(await load(face, { decoder: second }))).toEqual(bold);
  expect(second).toHaveBeenCalledTimes(2);
  release();
  expect(new Uint8Array(await bound)).toEqual(bold);
  expect(first).toHaveBeenCalledTimes(2);
  expect(decode).toHaveBeenCalledTimes(2);
});

test('bundled and CDN provider factories carry their host runtime and keep unbound decoding standalone', async () => {
  const { createFontProvider, loadBundledFontBytes, resolveMetricCompatFace } = await import('./index');
  const { createFontProvider: cdn } = await import('./cdn');
  const { default: manifest } = await import('../../fonts-latin/assets/manifest.json');
  const face = resolveMetricCompatFace('Calibri', false, false)!;
  const entry = manifest.faces.find(entry => entry.file === face.file)!;
  const requested: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requested.push(String(input));
    const encoded = new Uint8Array(entry.encodedLength);
    encoded[0] = 1;
    return new Response(encoded);
  }) as typeof fetch;
  const failure = new Error('linked decoder rejected face');
  const decoder = mock<FontDecoder>(async () => { throw failure; });
  const runtime = { decoder };
  const options = { baseUrl: `https://provider-runtime-${++sequence}.example/` };
  for (const factory of [createFontProvider, cdn]) {
    await expect(factory(options, runtime).resolve('Calibri', false, false)!()).rejects.toBe(failure);
  }
  await expect(loadBundledFontBytes(face, options, runtime)).rejects.toBe(failure);
  expect(decoder).toHaveBeenCalledTimes(3);
  expect(decode).not.toHaveBeenCalled();
  expect(requested.filter(url => url.endsWith('fontpack.wasm'))).toEqual([]);
  await expect(createFontProvider(options).resolve('Calibri', false, false)!()).rejects.toThrow();
  expect(decode).toHaveBeenCalledTimes(1);
  expect(decoder).toHaveBeenCalledTimes(3);
});

import { expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import pack from '../../fonts-cjk/assets/manifest.json';
import { createFontProvider as cdn } from './cdn';
import { decodeFontFace, type FontDecoder } from './decode';
import { createFontProvider } from './index';
import { openProgressiveFace, sliceSourceOf, type SliceLocation, type SliceRef, type SliceTransport } from './progressive';
import { openSlicedFace } from './progressive-entry';
import { chunkEntries, parseSkeleton } from './slices';

const assets = new URL('../../fonts-cjk/assets/', import.meta.url);
const sha = (bytes: ArrayBuffer) => createHash('sha256').update(new Uint8Array(bytes)).digest('hex');
const read = async (file: string) => Bun.file(new URL(file, assets)).arrayBuffer();
const decode = async (file: string) => new Uint8Array((await decodeFontFace(await read(file))).font);

interface Index {
  byteLength: number;
  sha256: string;
  skeleton: { file: string };
  core: { file: string };
  chunks: Array<{ id: number; file: string }>;
}

function transport(reads: string[], tamper?: (file: string) => boolean): SliceTransport {
  return {
    async read(file) {
      reads.push(file);
      const bytes = await read(file);
      return tamper?.(file) ? bytes.slice(0, bytes.byteLength - 1) : bytes;
    },
    async decode(encoded, maxOutput) {
      return (await decodeFontFace(encoded, undefined, { maxOutput })).font;
    },
  };
}

/** Serves `file` with one byte flipped after decoding, so only its SHA-256 can tell. */
function flipping(file: string): SliceTransport {
  const served = new WeakSet<ArrayBuffer>();
  return {
    async read(name) {
      const bytes = new Uint8Array(await read(name));
      if (name === file && name.endsWith('.json')) bytes[bytes.indexOf(0x30)] = 0x31;
      if (name === file) served.add(bytes.buffer);
      return bytes.buffer;
    },
    async decode(encoded, maxOutput) {
      const font = new Uint8Array((await decodeFontFace(encoded, undefined, { maxOutput })).font);
      if (served.has(encoded)) font[font.length >> 1] ^= 1;
      return font.buffer;
    },
  };
}

const jp = pack.faces.find((face) => face.file === 'NotoSansJP-Regular.otf.zst')!;
const jpRef = jp.slices as SliceRef;

test('every sliced face reassembles from all its chunks to its recorded SHA-256 with the whole face tables', async () => {
  for (const face of pack.faces) {
    const ref = face.slices as SliceRef;
    const index = JSON.parse(await Bun.file(new URL(ref.index, assets)).text()) as Index;
    const skeleton = parseSkeleton(await decode(index.skeleton.file));
    const outlines = new Map<number, Uint8Array>();
    for (const file of [index.core.file, ...index.chunks.map((chunk) => chunk.file)]) {
      for (const [glyph, outline] of chunkEntries(await decode(file))) outlines.set(glyph, outline);
    }
    const font = skeleton.reassemble(outlines);
    expect(font.byteLength).toBe(index.byteLength);
    expect(sha(font)).toBe(index.sha256);
    expect(skeleton.chunkOfCodepoint(0x20)).toBe(0);
    expect(skeleton.chunkOfGlyph(skeleton.numGlyphs)).toBeUndefined();
  }
});

test('a progressive face loads index, skeleton and core, then only the chunks of the glyphs asked for', async () => {
  const reads: string[] = [];
  const face = await openProgressiveFace(jpRef, transport(reads));
  const index = JSON.parse(await Bun.file(new URL(jpRef.index, assets)).text()) as Index;
  expect(reads.sort()).toEqual([jpRef.index, index.core.file, index.skeleton.file].sort());
  expect(face.revision).toBe(1);
  expect(face.complete()).toBe(false);
  const skeleton = parseSkeleton(await decode(index.skeleton.file));
  const glyphOf = (chunk: number) => {
    for (let glyph = 0; glyph < skeleton.numGlyphs; glyph += 1) {
      if (skeleton.chunkOfGlyph(glyph) === chunk) return glyph;
    }
    throw new Error(`no glyph in chunk ${chunk}`);
  };
  const updates: number[] = [];
  face.onUpdate((revision) => updates.push(revision));

  reads.length = 0;
  expect(await face.ensureGlyphs([glyphOf(0)])).toBe(1);
  expect(reads).toEqual([]);

  const [first, second] = await Promise.all([face.ensureGlyphs([glyphOf(40)]), face.ensureGlyphs([glyphOf(40), glyphOf(41)])]);
  expect(reads.sort()).toEqual([index.chunks[39]!.file, index.chunks[40]!.file].sort());
  expect(Math.max(first, second)).toBe(face.revision);
  expect(updates.length).toBeGreaterThan(0);
  expect(updates[updates.length - 1]).toBe(face.revision);

  const bytes = face.bytes();
  expect(face.bytes()).not.toBe(bytes);
  expect(face.bytes()).toEqual(bytes);
  expect(bytes.byteLength % 4).toBe(0);
  expect(new DataView(bytes).getUint32(0)).toBe(0x4f54544f);

  reads.length = 0;
  let codepoint = 0;
  while (skeleton.chunkOfCodepoint(codepoint) !== 60) codepoint += 1;
  face.prefetch(String.fromCodePoint(codepoint));
  await new Promise<void>((resolve) => {
    const stop = face.onUpdate(() => {
      stop();
      resolve();
    });
  });
  expect(reads).toEqual([index.chunks[59]!.file]);

  await face.ensureGlyphs(Array.from({ length: skeleton.numGlyphs }, (_, glyph) => glyph));
  expect(face.complete()).toBe(true);
  expect(sha(face.bytes())).toBe(index.sha256);
});

test('a glyph whose chunk loaded during another pending request resolves with a revision that draws it', async () => {
  const index = JSON.parse(await Bun.file(new URL(jpRef.index, assets)).text()) as Index;
  const slow = index.chunks[41]!.file;
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const fast: string[] = [];
  const base = transport([]);
  const face = await openProgressiveFace(jpRef, {
    read: async (file) => {
      if (file === slow) await gate;
      const bytes = await base.read(file);
      if (file === index.chunks[40]!.file) fast.push(file);
      return bytes;
    },
    decode: base.decode,
  });
  const skeleton = parseSkeleton(await decode(index.skeleton.file));
  const glyphIn = (chunk: number) => {
    let glyph = 0;
    while (skeleton.chunkOfGlyph(glyph) !== chunk) glyph += 1;
    return glyph;
  };
  const coreOnly = face.bytes();
  const both = face.ensureGlyphs([glyphIn(41), glyphIn(42)]);
  for (let tick = 0; tick < 200 && fast.length === 0; tick += 1) await new Promise((resolve) => setTimeout(resolve, 1));
  for (let tick = 0; tick < 20; tick += 1) await new Promise((resolve) => setTimeout(resolve, 1));
  const revision = await face.ensureGlyphs([glyphIn(41)]);
  expect(revision).toBeGreaterThan(1);
  expect(face.revision).toBe(revision);
  expect(face.bytes()).not.toEqual(coreOnly);
  release();
  expect(await both).toBe(face.revision);
});

test('a listener that throws neither stops the others nor rejects the load', async () => {
  const index = JSON.parse(await Bun.file(new URL(jpRef.index, assets)).text()) as Index;
  const face = await openProgressiveFace(jpRef, transport([]));
  const skeleton = parseSkeleton(await decode(index.skeleton.file));
  let glyph = 0;
  while (skeleton.chunkOfGlyph(glyph) !== 5) glyph += 1;
  const seen: number[] = [];
  face.onUpdate(() => {
    throw new Error('listener failed');
  });
  face.onUpdate((revision) => seen.push(revision));
  expect(await face.ensureGlyphs([glyph])).toBe(2);
  expect(seen).toEqual([2]);
});

test('a chunk that fails its checks rejects and loads on a later request', async () => {
  const reads: string[] = [];
  const index = JSON.parse(await Bun.file(new URL(jpRef.index, assets)).text()) as Index;
  const broken = index.chunks[9]!.file;
  let tampering = true;
  const face = await openProgressiveFace(jpRef, transport(reads, (file) => tampering && file === broken));
  const skeleton = parseSkeleton(await decode(index.skeleton.file));
  let glyph = 0;
  while (skeleton.chunkOfGlyph(glyph) !== 10) glyph += 1;
  await expect(face.ensureGlyphs([glyph])).rejects.toThrow('encoded bytes');
  expect(face.revision).toBe(1);
  tampering = false;
  expect(await face.ensureGlyphs([glyph])).toBe(2);
  expect(reads.filter((file) => file === broken)).toHaveLength(2);
});

test('opening rejects an index whose SHA-256 differs from the manifest pointer', async () => {
  await expect(openProgressiveFace({ ...jpRef, indexSha256: '0'.repeat(64) }, transport([])))
    .rejects.toThrow('SHA-256 mismatch');
});

test('a served index, skeleton, core or chunk changed at its recorded length is refused by its SHA-256', async () => {
  const index = JSON.parse(await Bun.file(new URL(jpRef.index, assets)).text()) as Index;
  await expect(openProgressiveFace(jpRef, flipping(jpRef.index))).rejects.toThrow(`Font slice index ${jpRef.index} SHA-256 mismatch`);
  for (const file of [index.skeleton.file, index.core.file]) {
    await expect(openProgressiveFace(jpRef, flipping(file))).rejects.toThrow(`Font slice ${file} SHA-256 mismatch`);
  }
  const chunk = index.chunks[9]!.file;
  const face = await openProgressiveFace(jpRef, flipping(chunk));
  const skeleton = parseSkeleton(await decode(index.skeleton.file));
  let glyph = 0;
  while (skeleton.chunkOfGlyph(glyph) !== 10) glyph += 1;
  await expect(face.ensureGlyphs([glyph])).rejects.toThrow(`Font slice ${chunk} SHA-256 mismatch`);
  expect(face.revision).toBe(1);
});

test('the progressive entry reaches no module that reads import.meta, so classic workers can load it', async () => {
  const build = await Bun.build({ entrypoints: [new URL('./progressive-entry.ts', import.meta.url).pathname], target: 'browser' });
  expect(build.success).toBe(true);
  expect(await build.outputs[0]!.text()).not.toContain('import.meta');
});

test('providers locate only sliced faces, keep whole-face loading, and share one opened face', async () => {
  const decoder: FontDecoder = (bytes, options) => decodeFontFace(bytes, options.dict, { codec: options.codec, maxOutput: options.maxOutput });
  const runtime = { decoder };
  for (const provider of [createFontProvider({ baseUrl: assets }, runtime), cdn({ baseUrl: assets }, runtime)]) {
    const sliced = provider.resolveScriptFallback('cjk-jp', false, false)!;
    expect(sliced.slices?.ref).toEqual(jpRef);
    expect(sliced.slices?.decoder).toBe(decoder);
    expect(sha(await sliced())).toBe(jp.sha256);
    expect(provider.resolve('Arial', false, false)?.slices).toBeUndefined();
  }
  const slices = createFontProvider({ baseUrl: assets }, runtime).resolveScriptFallback('cjk-jp', false, false)!.slices!;
  const face = await openSlicedFace(slices);
  expect(face.revision).toBeGreaterThanOrEqual(1);
  expect(await openSlicedFace(slices)).toBe(face);
  const failing = { ...slices, base: 'https://failing.example/', url: async () => { throw new Error('no such file'); } };
  await expect(openSlicedFace(failing)).rejects.toThrow('no such file');
  await expect(openSlicedFace({ ...failing, url: slices.url })).resolves.toBeDefined();
});

test('a face opened with file URLs names every slice file and the fetch timeout for another realm', async () => {
  const index = JSON.parse(await Bun.file(new URL(jpRef.index, assets)).text()) as Index;
  const resolved = (file: string) => new URL(file, 'https://fonts.example/cjk/');
  const face = await openProgressiveFace(jpRef, { ...transport([]), url: async (file) => resolved(file), timeoutMs: 30_000 });
  const source = await sliceSourceOf(face)!;
  expect(source.ref).toEqual(jpRef);
  expect(source.timeoutMs).toBe(30_000);
  const files = [jpRef.index, index.skeleton.file, index.core.file, ...index.chunks.map((chunk) => chunk.file)];
  expect(Object.keys(source.urls).sort()).toEqual(files.sort());
  expect(source.urls[index.chunks[0]!.file]).toBe(resolved(index.chunks[0]!.file).href);
  expect(sliceSourceOf(face)).toBe(sliceSourceOf(face));
  expect(sliceSourceOf(await openProgressiveFace(jpRef, transport([])))).toBeUndefined();
});

test('a location without file URLs cannot open, so callers fall back to the whole face', async () => {
  const { slices } = createFontProvider({ baseUrl: assets }).resolveScriptFallback('cjk-jp', false, false)!;
  const { url: _url, ...withoutUrl } = slices!;
  await expect(openSlicedFace({ ...withoutUrl, base: 'https://no-url.example/' } as SliceLocation)).rejects.toThrow();
});

test('retain: false opens a fresh face each time instead of keeping one cached', async () => {
  const decoder: FontDecoder = (bytes, options) => decodeFontFace(bytes, options.dict, { codec: options.codec, maxOutput: options.maxOutput });
  const kept = createFontProvider({ baseUrl: assets }, { decoder }).resolveScriptFallback('cjk-jp', false, false)!.slices!;
  const dropped = createFontProvider({ baseUrl: assets }, { decoder: (bytes, options) => decoder(bytes, options), retain: false })
    .resolveScriptFallback('cjk-jp', false, false)!.slices!;
  expect(dropped.retain).toBe(false);
  const first = await openSlicedFace(dropped);
  expect(await openSlicedFace(dropped)).not.toBe(first);
  expect(await openSlicedFace(kept)).toBe(await openSlicedFace(kept));
});

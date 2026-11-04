import { afterAll, beforeEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { decodeFontFace } from '../../../fonts/src/decode';
import { glyphsByFont, ResidentProgressiveFonts, type ProgressiveFontSpec } from './residentProgressiveFonts';

const assets = new URL('../../../fonts-cjk/assets/', import.meta.url);
const manifest = JSON.parse(readFileSync(new URL('manifest.json', assets), 'utf8')) as {
  faces: Array<{ file: string; slices?: ProgressiveFontSpec['source']['ref'] }>;
};
const ref = manifest.faces.find((face) => face.file === 'NotoSansJP-Regular.otf.zst')!.slices!;
const index = JSON.parse(readFileSync(new URL(ref.index, assets), 'utf8')) as {
  skeleton: { file: string };
  core: { file: string };
  chunks: Array<{ id: number; file: string }>;
};
const chunkFile = (id: number) => index.chunks.find((chunk) => chunk.id === id)!.file;

const hits = new Map<string, number>();
const failing = new Set<string>();
let gate: Promise<void> | null = null;
const server = Bun.serve({
  port: 0,
  async fetch(request) {
    const file = decodeURIComponent(new URL(request.url).pathname.slice(1));
    hits.set(file, (hits.get(file) ?? 0) + 1);
    if (failing.delete(file)) return new Response('unavailable', { status: 503 });
    if (gate && index.chunks.some((chunk) => chunk.file === file)) await gate;
    return new Response(Bun.file(new URL(file, assets)));
  },
});
afterAll(() => server.stop(true));
beforeEach(() => {
  hits.clear();
  failing.clear();
  gate = null;
});

const files = [ref.index, index.skeleton.file, index.core.file, ...index.chunks.map((chunk) => chunk.file)];
const spec: ProgressiveFontSpec = {
  slot: 0,
  source: { ref, urls: Object.fromEntries(files.map((file) => [file, new URL(file, server.url).href])), timeoutMs: 5_000 },
};
const decode = async (encoded: ArrayBuffer, maxOutput: number) => (await decodeFontFace(encoded, undefined, { maxOutput })).font;

const skeleton = new DataView(await decode(await Bun.file(new URL(index.skeleton.file, assets)).arrayBuffer(), 64 << 20));
const glyphIn = (chunk: number) => {
  const glyphs = skeleton.getUint32(8);
  for (let glyph = 0; glyph < glyphs; glyph += 1) if (skeleton.getUint16(32 + glyph * 2) === chunk) return glyph;
  throw new Error(`no glyph in chunk ${chunk}`);
};

function fonts(engine: (fontId: number, bytes: Uint8Array) => void = () => {},
  specs: ProgressiveFontSpec[] = [spec], ids: number[] = [10, 11], substituteOf: (number | undefined)[] = [undefined, 0]) {
  const logs: string[] = [];
  const updates: Array<{ fontId: number; length: number }> = [];
  const resident = new ResidentProgressiveFonts(specs, ids, substituteOf, {
    updateFontOutlines(fontId, bytes) {
      engine(fontId, bytes);
      updates.push({ fontId, length: bytes.byteLength });
    },
  }, decode, (message) => logs.push(message));
  return { resident, logs, updates };
}

test('a page walk collects the glyph ids of every run, including nested regions', () => {
  const page = {
    items: [{ kind: 'glyphRun', fontId: 10, glyphs: [{ id: 5 }, { id: 6 }] }],
    regions: [{ notes: [{ kind: 'glyphRun', fontId: 11, glyphs: [{ id: 7 }] }] }],
  };
  expect(glyphsByFont(page)).toEqual(new Map([[10, new Set([5, 6])], [11, new Set([7])]]));
});

test('outlines a frame draws install before it paints, one fetch per chunk across concurrent frames and substitutes', async () => {
  const { resident, updates } = fonts();
  const glyph = glyphIn(41);
  const [first, second] = await Promise.all([
    resident.ensure(new Map([[10, new Set([glyph])]])),
    resident.ensure(new Map([[11, new Set([glyph])]])),
  ]);
  expect(new Set([...first, ...second])).toEqual(new Set([10]));
  expect(updates.map((update) => update.fontId)).toEqual([10]);
  expect(hits.get(chunkFile(41))).toBe(1);
  expect([ref.index, index.skeleton.file, index.core.file].map((file) => hits.get(file))).toEqual([1, 1, 1]);
  expect(resident.fontIdsOf(new Set([10]))).toEqual(new Set([10, 11]));

  expect(await resident.ensure(new Map([[10, new Set([glyph, glyphIn(0)])]]))).toEqual(new Set());
  expect(updates).toHaveLength(1);
  expect(hits.get(chunkFile(41))).toBe(1);
  expect(await resident.ensure(new Map([[99, new Set([glyph])]]))).toEqual(new Set());
});

test('a frame stops waiting at the guard, logs once, and repaints when the outlines arrive', async () => {
  const { resident, logs, updates } = fonts();
  let release!: () => void;
  gate = new Promise((resolve) => {
    release = resolve;
  });
  let repaints = 0;
  resident.onLateInstall(() => {
    repaints += 1;
  });
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(50)])]]), 30)).toEqual(new Set());
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(51)])]]), 30)).toEqual(new Set());
  expect(logs.filter((message) => message.includes('still loading after 30 ms'))).toHaveLength(1);
  expect(updates).toHaveLength(0);
  release();
  await resident.whenInstalled();
  expect(repaints).toBeGreaterThanOrEqual(1);
  expect(resident.lateInstalls).toBe(repaints);
  expect(updates.length).toBeGreaterThanOrEqual(1);
  expect([chunkFile(50), chunkFile(51)].map((file) => hits.get(file))).toEqual([1, 1]);
});

test('a chunk that fails to load is fetched again', async () => {
  const { resident, updates, logs } = fonts();
  failing.add(chunkFile(60));
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(60)])]]))).toEqual(new Set([10]));
  expect(hits.get(chunkFile(60))).toBe(2);
  expect(updates).toHaveLength(1);
  expect(logs).toEqual([]);
});

test('an engine refusal is reported for a whole-face reload and is never retried with sliced outlines', async () => {
  let calls = 0;
  const { resident, logs } = fonts(() => {
    calls += 1;
    throw new Error('tables differ');
  });
  const refused: number[] = [];
  resident.onRefused((fontId) => refused.push(fontId));
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(70)])]]))).toEqual(new Set());
  expect(await resident.ensure(new Map([[11, new Set([glyphIn(71)])]]))).toEqual(new Set());
  expect(refused).toEqual([10]);
  expect(calls).toBe(1);
  expect(logs.some((message) => message.includes('refused the outlines') && message.includes('tables differ'))).toBe(true);
  expect(hits.get(chunkFile(71))).toBeUndefined();
});

test('slots registered with the same sliced face share one outline state and install to every slot', async () => {
  const { resident, updates } = fonts(undefined, [spec, { ...spec, slot: 1 }], [10, 20], [undefined, undefined]);
  const [first, second] = [glyphIn(41), glyphIn(42)];
  await Promise.all([
    resident.ensure(new Map([[10, new Set([first])]])),
    resident.ensure(new Map([[20, new Set([second])]])),
  ]);
  expect(hits.get(ref.index)).toBe(1);
  expect([chunkFile(41), chunkFile(42)].map((file) => hits.get(file))).toEqual([1, 1]);
  const last = new Map(updates.map((update) => [update.fontId, update.length]));
  expect([...last.keys()].sort()).toEqual([10, 20]);
  expect(last.get(10)).toBe(last.get(20));
  expect(updates.at(-1)!.length).toBe(Math.max(...updates.map((update) => update.length)));
  expect(resident.fontIdsOf(new Set([10]))).toEqual(new Set([10, 20]));
  expect(resident.size).toBe(1);
});

test('fonts that were replaced never install outlines that arrive later', async () => {
  const { resident, updates } = fonts();
  let release!: () => void;
  gate = new Promise((resolve) => {
    release = resolve;
  });
  let repaints = 0;
  resident.onLateInstall(() => {
    repaints += 1;
  });
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(52)])]]), 20)).toEqual(new Set());
  resident.dispose();
  release();
  await resident.whenInstalled();
  expect(updates).toEqual([]);
  expect(repaints).toBe(0);
  expect(await resident.ensure(new Map([[10, new Set([glyphIn(53)])]]))).toEqual(new Set());
  expect(hits.get(chunkFile(53))).toBeUndefined();
});

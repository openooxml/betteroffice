import { afterEach, beforeEach, expect, test } from 'bun:test';
import { prepareTextFamilies } from '../render/textFamilies';
import { createSlicedFontFaces, slicedFacesSupported, type SlicedFontFace } from './commandSlicedFonts';

let added: string[];
let removed: string[];
const originals = new Map<string, PropertyDescriptor | undefined>();
function set(key: string, value: unknown) {
  originals.set(key, Object.getOwnPropertyDescriptor(globalThis, key));
  Object.defineProperty(globalThis, key, { configurable: true, value });
}
beforeEach(() => {
  added = []; removed = [];
  set('fonts', { add: (face: { family: string }) => added.push(face.family), delete: (face: { family: string }) => removed.push(face.family) });
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { return this; }
  });
});
afterEach(() => {
  for (const [key, descriptor] of originals) {
    if (descriptor) Object.defineProperty(globalThis, key, descriptor);
    else Reflect.deleteProperty(globalThis, key);
  }
  originals.clear();
});

function fakeFace(fail: () => boolean = () => false): SlicedFontFace & { asked: string[] } {
  let revision = 1;
  const seen = new Set<number>();
  const chunksOf = (text: string) => [...`${text}${text.normalize('NFC')}`].filter((character) => character !== '\n')
    .map((character) => character.codePointAt(0)!);
  return {
    asked: [],
    get revision() { return revision; },
    bytes: () => new ArrayBuffer(8),
    prefetch() {},
    chunksOf,
    loadedChunks: () => seen,
    async ensureText(text) {
      this.asked.push(text);
      if (fail()) throw new Error('offline');
      const fresh = chunksOf(text).filter((chunk) => !seen.has(chunk));
      if (fresh.length) {
        for (const chunk of fresh) seen.add(chunk);
        revision += 1;
      }
      return revision;
    },
  };
}

test.each([
  ['Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36', true],
  ['Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) HeadlessChrome/149.0.0.0 Safari/537.36', true],
  ['Mozilla/5.0 (X11; Linux x86_64; rv:151.0) Gecko/20100101 Firefox/151.0', true],
  ['Mozilla/5.0 (Macintosh; Intel Mac OS X 15_0) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.0 Safari/605.1.15', false],
  ['Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/149.0 Mobile/15E148 Safari/604.1', false],
  ['Mozilla/5.0 (iPad; CPU OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) FxiOS/151.0 Mobile/15E148 Safari/605.1.15', false],
  ['', false],
])('sliced faces are used only in verified engines: %s', (agent, supported) => {
  expect(slicedFacesSupported(agent)).toBe(supported);
});

test('a family that is not sliced draws under its own name without loading anything', async () => {
  const sliced = createSlicedFontFaces();
  expect(await sliced.coverage.prepare('Arial', 'text')).toBe('Arial');
  expect(added).toEqual([]);
});

test('each revision installs under its own family, only when the text needed new outlines, and old ones are removed', async () => {
  const sliced = createSlicedFontFaces();
  const face = fakeFace();
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  expect(await sliced.coverage.prepare('noto sans jp', '日')).toBe('Noto Sans JP r2');
  expect(await sliced.coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP r2');
  expect(added).toEqual(['Noto Sans JP r2']);
  expect(await sliced.coverage.prepare('Noto Sans JP', '本')).toBe('Noto Sans JP r3');
  expect(await sliced.coverage.prepare('Noto Sans JP', '語')).toBe('Noto Sans JP r4');
  expect(added).toEqual(['Noto Sans JP r2', 'Noto Sans JP r3', 'Noto Sans JP r4']);
  expect(removed).toEqual(['Noto Sans JP r2']);
  sliced.dispose();
  expect(removed).toEqual(['Noto Sans JP r2', 'Noto Sans JP r3', 'Noto Sans JP r4']);
});

test('a chunk that cannot load gives way to the whole face, so nothing draws blank', async () => {
  const sliced = createSlicedFontFaces();
  let wholeLoads = 0;
  sliced.add('Noto Sans JP', fakeFace(() => true), async () => { wholeLoads += 1; return new ArrayBuffer(16); });
  expect(await sliced.coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP whole');
  expect(await sliced.coverage.prepare('Noto Sans JP', '本')).toBe('Noto Sans JP whole');
  expect(wholeLoads).toBe(1);
  expect((await sliced.read('Noto Sans JP', '語', null))?.bytes?.byteLength).toBe(16);
});

test('with neither chunks nor the whole face, the real family name is drawn, which falls back rather than blanks', async () => {
  const sliced = createSlicedFontFaces();
  sliced.add('Noto Sans JP', fakeFace(() => true), async () => { throw new Error('offline'); });
  expect(await sliced.coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP');
  expect(added).toEqual([]);
});

test('another realm gets bytes only when its revision is out of date', async () => {
  const sliced = createSlicedFontFaces();
  sliced.add('Noto Sans JP', fakeFace(), async () => new ArrayBuffer(16));
  const first = await sliced.read('Noto Sans JP', '日', null);
  expect(first?.bytes?.byteLength).toBe(8);
  const again = await sliced.read('Noto Sans JP', '日', first!.revision);
  expect(again).toEqual({ revision: first!.revision });
  expect(await sliced.read('Arial', 'text', null)).toBeUndefined();
});

test('an idle fill installs one revision, and later paints of that text fetch and install nothing', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  await sliced.fill('日本語日本語');
  expect(face.asked).toEqual(['日本語']);
  expect(added).toEqual(['Noto Sans JP r2']);
  expect(await sliced.coverage.prepare('Noto Sans JP', '日本')).toBe('Noto Sans JP r2');
  expect(face.asked).toEqual(['日本語']);
  expect(added).toEqual(['Noto Sans JP r2']);
});

test('a paint during an idle fill waits only for its own outlines, never for the fill', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  const ensure = face.ensureText.bind(face);
  face.ensureText = async (text) => { if (text === '語') await held; return ensure(text); };
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  const filling = sliced.fill('語');
  await new Promise((resolve) => setTimeout(resolve, 5));
  expect(await sliced.coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP r2');
  release();
  await filling;
});

test('an idle fill waits while paints are recent, and stops at disposal', async () => {
  const sliced = createSlicedFontFaces(40);
  const face = fakeFace();
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  await sliced.coverage.prepare('Noto Sans JP', '日');
  const asked = face.asked.length;
  const filling = sliced.fill('語');
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(face.asked).toHaveLength(asked);
  sliced.dispose();
  await filling;
  expect(face.asked).toHaveLength(asked);
});

test('paints that do not prepare their text get complete faces, and a face that cannot load rejects rather than paints blank', async () => {
  const sliced = createSlicedFontFaces();
  sliced.add('Noto Sans JP', fakeFace(), async () => new ArrayBuffer(16));
  expect((await sliced.wholeFaces(['Noto Sans JP'])).map(({ family, bytes }) => [family, bytes.byteLength])).toEqual([['Noto Sans JP', 16]]);
  sliced.add('Noto Sans KR', fakeFace(), async () => { throw new Error('offline'); });
  await expect(sliced.wholeFaces(['Noto Sans KR'])).rejects.toThrow('offline');
});

test('a revision only claims what its bytes hold, even when a fill finishes while it installs', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  let finishLoad!: () => void;
  const held = new Promise<void>((resolve) => { finishLoad = resolve; });
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { await held; return this; }
  });
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  const painting = sliced.coverage.prepare('Noto Sans JP', '日');
  await new Promise((resolve) => setTimeout(resolve, 0));
  const filling = sliced.fill('語');
  await new Promise((resolve) => setTimeout(resolve, 5));
  finishLoad();
  await painting;
  await filling;
  const family = await sliced.coverage.prepare('Noto Sans JP', '語');
  expect(family).not.toBe('Noto Sans JP r2');
});

test('one paint waiting on its chunk never holds up another paint', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  const ensure = face.ensureText.bind(face);
  face.ensureText = async (text) => { if (text === '語') await held; return ensure(text); };
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  const stalled = sliced.coverage.prepare('Noto Sans JP', '語');
  expect(await sliced.coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP r2');
  release();
  await stalled;
});

test('a font the host registers over one style stops paints slicing, and the other styles still export whole', async () => {
  const sliced = createSlicedFontFaces();
  sliced.add('Noto Sans JP', fakeFace(), async () => new Uint8Array([4]).buffer);
  await sliced.coverage.prepare('Noto Sans JP', '日');
  sliced.forget('noto sans jp', 7);
  expect(removed).toEqual(['Noto Sans JP r2']);
  expect(sliced.families()).toEqual([]);
  expect(await sliced.coverage.prepare('Noto Sans JP', '本')).toBe('Noto Sans JP');
  expect(sliced.coverage.latest!('Noto Sans JP')).toBeUndefined();
  expect(sliced.exportFamilies()).toEqual(['Noto Sans JP']);
  expect(sliced.keepsExportFace('Noto Sans JP', { id: 7 })).toBe(true);
  expect(sliced.keepsExportFace('Noto Sans JP', { id: 3 })).toBe(false);
  expect((await sliced.wholeFaces(['Noto Sans JP'])).map(({ bytes }) => [...new Uint8Array(bytes)])).toEqual([[4]]);
});

test('text whose chunk and whole face both fail keeps the original family, never an older revision', async () => {
  const sliced = createSlicedFontFaces();
  let offline = false;
  sliced.add('Noto Sans JP', fakeFace(() => offline), async () => { throw new Error('offline'); });
  await sliced.coverage.prepare('Noto Sans JP', '日');
  offline = true;
  const list = { primitives: [{ kind: 'textBox', paragraphs: [], lines: [{ runs: [{ text: '語', fontFamily: 'Noto Sans JP' }] }] }] };
  const families = await prepareTextFamilies(list as never, sliced.coverage);
  expect(families('Noto Sans JP')).toBe('Noto Sans JP');
});

test('a whole face still downloading never holds up another paint', async () => {
  const sliced = createSlicedFontFaces();
  const face = fakeFace();
  const ensure = face.ensureText.bind(face);
  face.ensureText = async (text) => { if (text === '語') throw new Error('offline'); return ensure(text); };
  let release!: (bytes: ArrayBuffer) => void;
  sliced.add('Noto Sans JP', face, () => new Promise((resolve) => { release = resolve; }));
  const stalled = sliced.coverage.prepare('Noto Sans JP', '語');
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(await sliced.coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP r2');
  release(new ArrayBuffer(16));
  await stalled;
});

test('callers asking at once download the whole face once, and a failure can be retried', async () => {
  const sliced = createSlicedFontFaces();
  let loads = 0;
  sliced.add('Noto Sans JP', fakeFace(), async () => {
    loads += 1;
    if (loads === 1) throw new Error('offline');
    return new ArrayBuffer(16);
  });
  await expect(sliced.whole('Noto Sans JP')!).rejects.toThrow('offline');
  await Promise.all([sliced.whole('Noto Sans JP'), sliced.wholeFaces(['Noto Sans JP']), sliced.whole('Noto Sans JP')]);
  expect(loads).toBe(2);
});

test('a new syllable made of jamo already seen is not treated as covered', async () => {
  const sliced = createSlicedFontFaces();
  const face = fakeFace();
  sliced.add('Noto Sans KR', face, async () => new ArrayBuffer(16));
  await sliced.coverage.prepare('Noto Sans KR', '\u1100\n\u1161\n\u11a9');
  face.asked.length = 0;
  await sliced.coverage.prepare('Noto Sans KR', '\u1100\u1161\u11a9');
  expect(face.asked).toEqual(['\u1100\u1161\u11a9']);
});

test('new text with combining marks asks the face, which alone knows what it composes into', async () => {
  const sliced = createSlicedFontFaces();
  const face = fakeFace();
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  await sliced.coverage.prepare('Noto Sans JP', 'A\u030a\u0301');
  face.asked.length = 0;
  await sliced.coverage.prepare('Noto Sans JP', 'A\u0301\u030a');
  expect(face.asked).toEqual(['A\u0301\u030a']);
});

test('repainting the same marked text never waits on a background parse', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { if (this.family.endsWith(' r3')) await held; return this; }
  });
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  expect(await sliced.coverage.prepare('Noto Sans JP', '\u304b\u3099')).toBe('Noto Sans JP r2');
  const filling = sliced.fill('語');
  await new Promise((resolve) => setTimeout(resolve, 5));
  face.asked.length = 0;
  expect(await sliced.coverage.prepare('Noto Sans JP', '\u304b\u3099')).toBe('Noto Sans JP r2');
  expect(face.asked).toEqual([]);
  release();
  await filling;
});

test('new right-to-left text asks the face, which alone knows what it mirrors into', async () => {
  const sliced = createSlicedFontFaces();
  const face = fakeFace();
  sliced.add('Noto Sans SC', face, async () => new ArrayBuffer(16));
  await sliced.coverage.prepare('Noto Sans SC', '\u202e\u2208\u202c');
  face.asked.length = 0;
  await sliced.coverage.prepare('Noto Sans SC', '\u202e\u2208\u2203\u202c');
  expect(face.asked).toEqual(['\u202e\u2208\u2203\u202c']);
});

test('a style that loaded whole under a sliced family keeps its own name and its export bytes', () => {
  const sliced = createSlicedFontFaces();
  sliced.add('Noto Sans JP', fakeFace(), async () => new ArrayBuffer(16));
  sliced.plain('noto sans jp', true, false);
  expect(sliced.coverage.covers!('Noto Sans JP', true, false)).toBe(false);
  expect(sliced.coverage.covers!('Noto Sans JP', false, false)).toBe(true);
  expect(sliced.keepsExportFace('Noto Sans JP', { id: 4, bold: true })).toBe(true);
  expect(sliced.keepsExportFace('Noto Sans JP', { id: 5 })).toBe(false);
  expect(sliced.plainStyles()).toEqual(['noto sans jp|1|0']);
});

test('a face registering after its whole bytes loaded can be upgraded on the spot, retired family included', async () => {
  const sliced = createSlicedFontFaces();
  const whole = new ArrayBuffer(16);
  sliced.add('Noto Sans KR', fakeFace(), async () => whole);
  expect(sliced.loadedWhole('Noto Sans KR')).toBeUndefined();
  await sliced.whole('Noto Sans KR');
  expect(sliced.loadedWhole('noto sans kr')).toBe(whole);
  sliced.forget('Noto Sans KR', 9);
  expect(sliced.loadedWhole('Noto Sans KR')).toBe(whole);
  expect(sliced.loadedWhole('Arial')).toBeUndefined();
});

test('a paint never waits for a background revision to finish parsing, and that revision then gives way', async () => {
  const sliced = createSlicedFontFaces(0);
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { if (this.family.endsWith(' r2')) await held; return this; }
  });
  sliced.add('Noto Sans JP', fakeFace(), async () => new ArrayBuffer(16));
  const filling = sliced.fill('語');
  await new Promise((resolve) => setTimeout(resolve, 5));
  expect(await sliced.coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP r3');
  release();
  await filling;
  expect(added).toEqual(['Noto Sans JP r3']);
  expect(await sliced.coverage.prepare('Noto Sans JP', '語')).toBe('Noto Sans JP r3');
});

test('two paints reaching the same revision parse it once', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  let parses = 0;
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { parses += 1; return this; }
  });
  face.ensureText = async () => 2;
  const revision = Object.getOwnPropertyDescriptor(face, 'revision')!;
  Object.defineProperty(face, 'revision', { ...revision, get: () => 2 });
  const families = await Promise.all([sliced.coverage.prepare('Noto Sans JP', '日'), sliced.coverage.prepare('Noto Sans JP', '語')]);
  expect(families).toEqual(['Noto Sans JP r2', 'Noto Sans JP r2']);
  expect(parses).toBe(1);
});

test('sliced font code stays off the opening path: nothing there imports it statically', async () => {
  const { readFile } = await import('node:fs/promises');
  const source = (path: string) => readFile(new URL(path, import.meta.url), 'utf8');
  const staticImport = (text: string, module: string) => new RegExp(`^import (?!type )[^;]*from '[^']*${module}'`, 'm').test(text);
  for (const path of ['./commandHost.ts', './commandDefaultFonts.ts', './workerRender.ts', '../render/canvas.ts', './host.ts']) {
    const text = await source(path);
    expect(staticImport(text, 'commandSlicedFonts')).toBe(false);
    expect(staticImport(text, 'textFamilies')).toBe(false);
  }
});

test('marked text needing only loaded chunks never waits on a background parse, even as a new string', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  set('FontFace', class {
    constructor(readonly family: string) {}
    async load() { if (this.family.endsWith(' r3')) await held; return this; }
  });
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  expect(await sliced.coverage.prepare('Noto Sans JP', '\u304b\u3099\nA')).toBe('Noto Sans JP r2');
  const filling = sliced.fill('語');
  await new Promise((resolve) => setTimeout(resolve, 5));
  face.asked.length = 0;
  expect(await sliced.coverage.prepare('Noto Sans JP', 'A\n\u304b\u3099')).toBe('Noto Sans JP r2');
  expect(face.asked).toEqual([]);
  release();
  await filling;
});

test('what a revision holds is a chunk set, so editing marked text retains nothing per string', async () => {
  const sliced = createSlicedFontFaces(0);
  const face = fakeFace();
  sliced.add('Noto Sans JP', face, async () => new ArrayBuffer(16));
  for (let edit = 0; edit < 50; edit += 1) await sliced.coverage.prepare('Noto Sans JP', `\u304b\u3099${'A'.repeat(edit)}`);
  expect(face.loadedChunks().size).toBeLessThan(6);
});


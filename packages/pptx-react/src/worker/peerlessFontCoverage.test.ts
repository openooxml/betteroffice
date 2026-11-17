import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeEach, expect, mock, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const { createPeerlessFontCoverage, createWholeFontInstaller } = await import('./peerlessFontCoverage');
import type { PresentationCommandClient } from '@betteroffice/pptx/worker';

let added: string[];
let removed: string[];
beforeEach(() => {
  added = []; removed = [];
  Object.defineProperty(document, 'fonts', { configurable: true, value: {
    add: (face: { family: string }) => added.push(face.family), delete: (face: { family: string }) => removed.push(face.family) } });
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: class {
    constructor(readonly family: string) {}
    async load() { return this; }
  } });
});

const sliced = { slicedFamilies: ['Noto Sans JP', 'MS Gothic'] } as PresentationCommandClient['state'];

test('a family the worker did not slice never asks it', async () => {
  const readCoverageFont = mock(async () => null);
  const coverage = createPeerlessFontCoverage({ readCoverageFont, state: sliced }, () => true);
  expect(await coverage.prepare('Calibri', 'text')).toBe('Calibri');
  expect(readCoverageFont).not.toHaveBeenCalled();
});

test('one paint asks once with all its text, and a stale revision installs under its own family', async () => {
  let revision = 1;
  const readCoverageFont = mock(async (_family: string, _text: string, have: number | null) => {
    revision += 1;
    return have === revision ? { revision } : { revision, bytes: new ArrayBuffer(8) };
  });
  const coverage = createPeerlessFontCoverage({ readCoverageFont, state: sliced }, () => true);
  const families = await Promise.all([coverage.prepare('Noto Sans JP', '日本'), coverage.prepare('Noto Sans JP', '語')]);
  const [first] = families;
  expect(first).toMatch(/^Noto Sans JP s\d+ r2$/);
  expect(families).toEqual([first, first]);
  expect(readCoverageFont.mock.calls).toEqual([['Noto Sans JP', '日本\n語', null]]);
  expect(added).toEqual([first]);
  await coverage.prepare('Noto Sans JP', 'a');
  await coverage.prepare('Noto Sans JP', 'b');
  expect(added.map((family) => family.replace(/ s\d+/, ''))).toEqual(['Noto Sans JP r2', 'Noto Sans JP r3', 'Noto Sans JP r4']);
  expect(removed).toEqual([first!]);
  expect(coverage.latest!('Noto Sans JP')).toBe(added[2]);
});

test('a sliced family under a named font asks too, and a worker with no face draws under its own name', async () => {
  const readCoverageFont = mock(async (_family: string, _text: string, _have: number | null) => null);
  const coverage = createPeerlessFontCoverage({ readCoverageFont, state: sliced }, () => true);
  expect(await coverage.prepare('MS Gothic', 'abc')).toBe('MS Gothic');
  expect(readCoverageFont.mock.calls).toEqual([['MS Gothic', 'abc', null]]);
});

test('two editors on one page never share a revision family', async () => {
  const read = async () => ({ revision: 2, bytes: new ArrayBuffer(8) });
  const one = createPeerlessFontCoverage({ readCoverageFont: read, state: sliced }, () => true);
  const two = createPeerlessFontCoverage({ readCoverageFont: read, state: sliced }, () => true);
  expect(await one.prepare('Noto Sans JP', '日')).not.toBe(await two.prepare('Noto Sans JP', '語'));
});

test('a family the worker stops slicing drops its page faces, so nothing draws from them again', async () => {
  const state = { slicedFamilies: ['Noto Sans JP'] };
  const coverage = createPeerlessFontCoverage({ readCoverageFont: async () => ({ revision: 2, bytes: new ArrayBuffer(8) }), state } as never, () => true);
  const first = await coverage.prepare('Noto Sans JP', '日');
  state.slicedFamilies = [];
  expect(await coverage.prepare('Noto Sans JP', '日')).toBe('Noto Sans JP');
  expect(removed).toEqual([first]);
  expect(coverage.latest!('Noto Sans JP')).toBeUndefined();
});

test('the slide being edited never waits for a chunk only a thumbnail needs', async () => {
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  const coverage = createPeerlessFontCoverage({ state: sliced, readCoverageFont: async (_: string, text: string) => {
    if (text === '語') await held;
    return { revision: 2, bytes: new ArrayBuffer(8) };
  } } as never, () => true);
  const thumbnail = coverage.thumbnails.prepare('Noto Sans JP', '語');
  expect(await coverage.prepare('Noto Sans JP', '日')).toMatch(/ r2$/);
  release();
  await thumbnail;
});

test('a case variant of a sliced family is sliced too', async () => {
  const asked: string[] = [];
  const coverage = createPeerlessFontCoverage({ state: sliced, readCoverageFont: async (family: string) => {
    asked.push(family);
    return { revision: 2, bytes: new ArrayBuffer(8) };
  } } as never, () => true);
  expect(await coverage.prepare('noto sans jp', '日')).toMatch(/ r2$/);
  expect(asked).toEqual(['noto sans jp']);
});

test('whole faces load once for callers at the same moment, and not again once installed', async () => {
  const reads: string[][] = [];
  const installs: string[][] = [];
  const ensure = createWholeFontInstaller(async (have) => {
    reads.push(have);
    return have.includes('Noto Sans JP') ? [] : [{ family: 'Noto Sans JP' }];
  }, async (faces) => { installs.push(faces.map((face) => face.family)); return faces; }, () => false);
  await Promise.all([ensure(), ensure(), ensure()]);
  expect(installs).toEqual([['Noto Sans JP']]);
  await ensure();
  expect(reads[0]).toEqual([]);
  expect(reads.slice(1).every((have) => have.includes('Noto Sans JP'))).toBe(true);
  expect(installs).toHaveLength(1);
});

test('a whole face arriving after the host supplied its family is dropped', async () => {
  const supplied = new Set<string>();
  const installs: string[] = [];
  const ensure = createWholeFontInstaller(async () => {
    supplied.add('Noto Sans JP');
    return [{ family: 'Noto Sans JP' }];
  }, async (faces) => { installs.push(...faces.map((face) => face.family)); return faces; }, (family) => supplied.has(family));
  await ensure();
  expect(installs).toEqual([]);
});

test('toolbar PNG export paints through coverage, so it never depends on a whole face loading', async () => {
  const { readFile } = await import('node:fs/promises');
  const source = await readFile(new URL('./peerlessFrameSource.ts', import.meta.url), 'utf8');
  const exporting = source.slice(source.indexOf('export async function peerlessExportPng'));
  expect(exporting).toContain('fontCoverage: run.fontCoverage');
  expect(exporting).not.toContain('ensureWholeFonts');
});

test('a right-to-left page asks for its text as right-to-left, since a page canvas inherits direction', async () => {
  const asked: string[] = [];
  const coverage = createPeerlessFontCoverage({ state: sliced, readCoverageFont: async (_: string, text: string) => {
    asked.push(text);
    return { revision: 2, bytes: new ArrayBuffer(8) };
  } } as never, () => true);
  document.documentElement.dir = 'rtl';
  try { await coverage.prepare('Noto Sans JP', '\u2208'); } finally { document.documentElement.dir = ''; }
  expect(asked).toEqual(['\u2208\u200f']);
});

test('a family that arrives during a whole-face load is installed before that load settles', async () => {
  const available = ['Noto Sans JP'];
  let release!: () => void;
  const held = new Promise<void>((resolve) => { release = resolve; });
  const installs: string[] = [];
  let reads = 0;
  const ensure = createWholeFontInstaller(async (have) => {
    const faces = available.filter((family) => !have.includes(family)).map((family) => ({ family }));
    reads += 1;
    if (reads === 1) await held;
    return faces;
  }, async (faces) => { installs.push(...faces.map((face) => face.family)); return faces; }, () => false);
  const first = ensure();
  await new Promise((resolve) => setTimeout(resolve, 0));
  available.push('Noto Sans SC');
  const second = ensure();
  release();
  await Promise.all([first, second]);
  expect(installs).toEqual(['Noto Sans JP', 'Noto Sans SC']);
});

test('a style that loaded whole under a sliced family is not covered, and the rest of the family is', () => {
  const state = { slicedFamilies: ['Noto Sans JP'], slicedPlainStyles: ['noto sans jp|1|0'] };
  const coverage = createPeerlessFontCoverage({ readCoverageFont: async () => null, state } as never, () => true);
  expect(coverage.covers!('Noto Sans JP', true, false)).toBe(false);
  expect(coverage.covers!('Noto Sans JP', false, false)).toBe(true);
  expect(coverage.thumbnails.covers!('Noto Sans JP', true, false)).toBe(false);
});

test('the page loads this module only for a session with sliced faces', async () => {
  const { readFile } = await import('node:fs/promises');
  const source = await readFile(new URL('./usePeerlessPresentation.ts', import.meta.url), 'utf8');
  expect(/^import (?!type )[^;]*from '[^']*peerlessFontCoverage'/m.test(source)).toBe(false);
  expect(source).toContain("props.experimentalSlicedFonts ? import('./peerlessFontCoverage') : undefined");
});


import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import * as fonts from '@betteroffice/fonts';
import * as progressive from '@betteroffice/fonts/progressive';
import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';
import { commandFontProvider, createCommandDefaultFonts, prefetchCoverageFaces } from './commandDefaultFonts';
import { createSlicedFontFaces } from './commandSlicedFonts';
import { decodeFontFace } from '../wasm/loader';
import { engineKind } from '../wasm/engine';

const restorers: (() => void)[] = [];

async function notoSansJP(): Promise<ArrayBuffer> {
  const packed = await readFile(new URL('../../../fonts-cjk/assets/NotoSansJP-Regular.otf.zst', import.meta.url));
  return (await decodeFontFace(packed, { codec: 'zstd' })).font;
}
const regular = { family: 'Calibri', bold: false, italic: false };
const bytes = new Uint8Array([1, 2, 3]).buffer;

function provider(
  load: (family: string, bold: boolean, italic: boolean) => (() => Promise<ArrayBuffer>) | undefined,
  script: (script: string, bold: boolean, italic: boolean) => (() => Promise<ArrayBuffer>) | undefined = () => undefined,
) {
  const resolve = mock(load);
  const resolveFamily = mock(() => undefined);
  const resolveScriptFallback = mock(script);
  const resolveLastResort = mock((_family: string, _bold: boolean, _italic: boolean, _office?: string) =>
    async (): Promise<ArrayBuffer> => { throw new Error('No substitute'); });
  const source = { resolve, resolveFamily, resolveScriptFallback, resolveLastResort };
  const imported = spyOn(commandFontProvider, 'load').mockResolvedValue(source);
  restorers.push(() => imported.mockRestore());
  return { imported, resolve, resolveFamily, resolveScriptFallback, resolveLastResort };
}

afterEach(() => { for (const restore of restorers.splice(0).reverse()) restore(); });

test('the manager decodes packed faces with the worker engine', async () => {
  const create = spyOn(fonts, 'createFontProvider');
  restorers.push(() => create.mockRestore());
  await commandFontProvider.load();
  const decoder = (create.mock.calls[0]![1] as { decoder: typeof decodeFontFace }).decoder;
  const packed = await readFile(new URL('../../../fonts-cjk/assets/NotoSansJP-Regular.otf.zst', import.meta.url));
  expect((await decoder(packed, { codec: 'zstd' })).font.byteLength).toBe((await notoSansJP()).byteLength);
});

test('coverage prefetch loads each face once; faces decode only after the session engine is ready, except large faces on wasm', async () => {
  let release!: () => void;
  const ready = new Promise<void>((resolve) => { release = resolve; });
  const load = mock(async () => bytes);
  const { resolveScriptFallback } = provider(() => undefined, (script) => script === 'cjk-jp' ? load : undefined);
  prefetchCoverageFaces(['cjk-jp'], ready, () => commandFontProvider.load());
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(resolveScriptFallback.mock.calls).toEqual([['cjk-jp', false, false]]);
  expect(load).toHaveBeenCalledTimes(1);
  restorers.splice(0).forEach((restore) => restore());
  const create = spyOn(fonts, 'createFontProvider');
  restorers.push(() => create.mockRestore());
  await commandFontProvider.load();
  const decoder = (create.mock.calls[0]![1] as { decoder: typeof decodeFontFace }).decoder;
  const large = await readFile(new URL('../../../fonts-cjk/assets/NotoSansJP-Regular.otf.zst', import.meta.url));
  let largeDecoded = false;
  const largeDecoding = decoder(large, { codec: 'zstd' }).then((face) => { largeDecoded = true; return face; });
  if (engineKind === 'wasm') await largeDecoding;
  const small = await readFile(new URL('../../../fonts-latin/assets/LiberationSans-Regular.ttf.zst', import.meta.url));
  let decoded = false;
  const decoding = decoder(small, { codec: 'woff2t+zstd' }).then((face) => { decoded = true; return face; });
  await new Promise((resolve) => setTimeout(resolve, 20));
  expect(decoded).toBe(false);
  expect(largeDecoded).toBe(engineKind === 'wasm');
  release();
  expect((await decoding).font.byteLength).toBeGreaterThan(0);
  expect((await largeDecoding).font.byteLength).toBeGreaterThan(0);
});

test('a prefetched coverage face and the later request share one decode', async () => {
  const original = fonts.createFontProvider;
  const decode = mock(decodeFontFace);
  const create = spyOn(fonts, 'createFontProvider').mockImplementation((options) => original(options, { decoder: decode }));
  restorers.push(() => create.mockRestore());
  const shared = commandFontProvider.load();
  prefetchCoverageFaces(['cjk-jp'], Promise.resolve(), () => shared);
  const source = await shared;
  const japanese = await source.resolveScriptFallback('cjk-jp', false, false)!();
  expect(japanese.byteLength).toBe((await notoSansJP()).byteLength);
  expect(decode).toHaveBeenCalledTimes(1);
});

test('coverage prefetch without scripts loads nothing', async () => {
  const { imported } = provider(() => undefined);
  prefetchCoverageFaces([], Promise.resolve(), () => commandFontProvider.load());
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(imported).not.toHaveBeenCalled();
});

test('app families exclude all styles without importing the manager', async () => {
  const source = provider(() => async () => bytes);
  const register = mock(() => {});
  const beforeRegister = mock(() => {});
  const fonts = createCommandDefaultFonts([{ family: ' calibri ' }], () => [regular, { ...regular, bold: true }],
    register, () => {}, undefined, beforeRegister);
  try {
    expect(await fonts.load()).toEqual([]);
    expect(source.imported).not.toHaveBeenCalled();
    expect(register).not.toHaveBeenCalled();
    expect(beforeRegister).not.toHaveBeenCalled();
  } finally { fonts.dispose(); }
});

test('defaults prepare once before the first registration across loads', async () => {
  provider(() => async () => bytes);
  const calls: string[] = [];
  let wanted = [regular];
  const fonts = createCommandDefaultFonts([], () => wanted, (face) => { calls.push(face.family); }, () => {},
    undefined, () => { calls.push('prepare'); });
  try {
    expect(calls).toEqual([]);
    await fonts.load();
    wanted = [{ ...regular, family: 'Cambria' }];
    await fonts.load();
    await fonts.load();
    expect(calls).toEqual(['prepare', 'Calibri', 'Cambria']);
  } finally { fonts.dispose(); }
});

test('unavailable and rejected defaults never prepare registration', async () => {
  provider((family) => family === 'Calibri' ? async () => bytes : undefined);
  const beforeRegister = mock(() => {});
  const register = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [regular, { ...regular, family: 'Unknown Family' }],
    register, () => {}, async () => [], beforeRegister);
  try {
    expect(await fonts.load()).toEqual([]);
    expect(register).not.toHaveBeenCalled();
    expect(beforeRegister).not.toHaveBeenCalled();
  } finally { fonts.dispose(); }
});

test('resolves requested traits through manager mappings and gives unknown families the PowerPoint last resort', async () => {
  const source = provider((family) => family === 'Calibri' ? async () => bytes : undefined);
  source.resolveLastResort.mockImplementation(() => async () => bytes);
  const register = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [{ ...regular, bold: true, italic: true },
    { ...regular, family: 'Unknown Family' }], register, () => {});
  try {
    const faces = await fonts.load();
    expect(source.resolve.mock.calls).toEqual([['Calibri', true, true], ['Unknown Family', false, false]]);
    expect(source.resolveFamily).toHaveBeenCalledWith('Unknown Family', false, false);
    expect(source.resolveLastResort.mock.calls).toEqual([['Unknown Family', false, false, 'powerpoint']]);
    expect(faces).toHaveLength(2);
    expect(faces[0]).toMatchObject({ family: 'Calibri', bold: true, italic: true });
    expect(faces[1]).toMatchObject({ family: 'Unknown Family', bold: false, italic: false });
    expect(faces[0].ownWeight).toBeUndefined();
    expect(faces.every((face) => !('bytes' in face))).toBe(true);
    expect(bytes.byteLength).toBe(3);
  } finally { fonts.dispose(); }
});

test('lighter members of bundled families load their regular face at their own weight; other unknown names take the last resort', async () => {
  const resolve = mock(() => undefined);
  const resolveLightVariant = mock((family: string, _bold: boolean, _italic: boolean) =>
    family === 'Calibri Light' ? async () => bytes : undefined);
  const resolveLastResort = mock((_family: string, _bold: boolean, _italic: boolean, _office?: string) => async () => bytes);
  const imported = spyOn(commandFontProvider, 'load').mockResolvedValue({ resolve, resolveFamily: () => undefined,
    resolveScriptFallback: () => undefined, resolveLastResort, resolveLightVariant });
  restorers.push(() => imported.mockRestore());
  const register = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [{ ...regular, family: 'Calibri Light', bold: true },
    { ...regular, family: 'BIZ UD明朝 Medium' }], register, () => {});
  try {
    const faces = await fonts.load();
    expect(resolveLightVariant.mock.calls).toEqual([['Calibri Light', true, false], ['BIZ UD明朝 Medium', false, false]]);
    expect(resolveLastResort.mock.calls).toEqual([['BIZ UD明朝 Medium', false, false, 'powerpoint']]);
    expect(faces.map((face) => [face.family, face.bold, face.ownWeight])).toEqual([['Calibri Light', true, true], ['BIZ UD明朝 Medium', false, undefined]]);
    expect(register).toHaveBeenCalledTimes(2);
  } finally { fonts.dispose(); }
});

test('failed, invalid and timed out faces settle once without late registration', async () => {
  let finish!: (value: ArrayBuffer) => void;
  const slow = new Promise<ArrayBuffer>((resolve) => { finish = resolve; });
  provider((family) => family === 'Slow' ? () => slow : family === 'Failure'
    ? async () => { throw new Error('Unavailable'); } : async () => bytes);
  const timeout = globalThis.setTimeout;
  const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: TimerHandler, delay?: number, ...args: unknown[]) =>
    timeout(callback, delay === 5000 ? 10 : delay, ...args)) as typeof setTimeout);
  restorers.push(() => timer.mockRestore());
  const register = mock((face: { family: string }) => { if (face.family === 'Invalid') throw new Error('Invalid font'); });
  const settled = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => ['Slow', 'Failure', 'Invalid', 'Calibri']
    .map((family) => ({ ...regular, family })), register, settled);
  try {
    expect(await fonts.load()).toMatchObject([{ family: 'Calibri' }]);
    expect(settled).toHaveBeenCalledTimes(1);
    finish(bytes);
    await Promise.resolve();
    expect(await fonts.load()).toEqual([]);
    expect(register.mock.calls.map(([face]) => face.family)).toEqual(['Invalid', 'Calibri']);
  } finally { fonts.dispose(); }
});

test('disposing an in-flight load settles without registering', async () => {
  provider(() => () => new Promise(() => {}));
  const register = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [regular], register, () => {});
  const loading = fonts.load();
  fonts.dispose();
  expect(await loading).toEqual([]);
  expect(register).not.toHaveBeenCalled();
});

test('default font plumbing stays out of all package entries', async () => {
  for (const entry of ['../index.ts', '../worker.ts', '../collaboration/index.ts', '../../../pptx-react/src/index.ts']) {
    const source = await readFile(new URL(entry, import.meta.url), 'utf8');
    expect(source).not.toMatch(/commandDefaultFonts|pptxDefaultFonts|presentationRequestedFontFaces|presentationEnableBrowserCoverageFallback|defaultFontInstaller|installDefaultFonts/);
  }
});

test('script requests keep traits, fall back to regular and deduplicate across families and concurrent loads', async () => {
  const arabic = new Uint8Array(await readFile(new URL('../../../fonts/assets/NotoSansArabic-Regular.ttf', import.meta.url))).buffer;
  const source = provider(() => undefined, (script, bold, italic) =>
    script === 'arabic' && !bold && !italic ? async () => arabic : undefined);
  const register = mock(() => {});
  const settled = mock(() => {});
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces);
  const requested = { ...regular, bold: true, italic: true, scripts: ['arabic', 'arabic'] };
  const fonts = createCommandDefaultFonts([{ family: 'Calibri' }], () => [requested,
    { ...requested, family: 'Unknown Family' }], register, settled, install);
  try {
    expect(fonts.hasPending()).toBe(true);
    const results = await Promise.all([fonts.load(), fonts.load()]);
    expect(source.resolve.mock.calls).toEqual([['Unknown Family', true, true]]);
    expect(source.resolveScriptFallback.mock.calls).toEqual([['arabic', true, true], ['arabic', false, false]]);
    expect(results.flat()).toMatchObject([{ family: 'Noto Sans Arabic', bold: false, italic: false, kind: 'coverage' }]);
    expect(install).toHaveBeenCalledTimes(1);
    expect(install.mock.calls[0][0][0].kind).toBe('coverage');
    expect(register).toHaveBeenCalledTimes(1);
    expect(settled).toHaveBeenCalledTimes(1);
    expect(fonts.hasPending()).toBe(false);
    expect(await fonts.load()).toEqual([]);
  } finally { fonts.dispose(); }
});

test('script collection requests each style separately even when no script pack is available', async () => {
  const source = provider(() => undefined);
  const fonts = createCommandDefaultFonts([], () => [
    { ...regular, family: 'BIZ UD明朝', bold: true, scripts: ['cjk-jp'] },
    { ...regular, family: 'Unknown Indic', italic: true, scripts: ['devanagari'] },
  ], () => {}, () => {});
  try {
    expect(await fonts.load()).toEqual([]);
    expect(source.resolveScriptFallback.mock.calls).toEqual([
      ['cjk-jp', true, false], ['cjk-jp', false, false],
      ['devanagari', false, true], ['devanagari', false, false],
    ]);
    expect(await fonts.load()).toEqual([]);
    expect(source.resolveScriptFallback).toHaveBeenCalledTimes(4);
  } finally { fonts.dispose(); }
});

test('styled Japanese requests share installation by the resolved regular face', async () => {
  const japanese = await notoSansJP();
  const source = provider(() => undefined, (script) => script === 'cjk-jp' ? async () => japanese : undefined);
  const register = mock((_face: DefaultFontFace) => {});
  const settled = mock(() => {});
  let finish!: () => void;
  let started!: () => void;
  const installing = new Promise<void>((resolve) => { finish = resolve; });
  const requested = new Promise<void>((resolve) => { started = resolve; });
  const install = mock(async (faces: readonly DefaultFontFace[]) => { started(); await installing; return faces; });
  const traits = [[false, false], [true, false], [false, true], [true, true]] as const;
  const fonts = createCommandDefaultFonts([{ family: 'Calibri' }], () => traits.map(([bold, italic]) =>
    ({ ...regular, bold, italic, scripts: ['cjk-jp'] })), register, settled, install);
  try {
    const first = fonts.load();
    const second = fonts.load();
    await requested;
    finish();
    expect((await Promise.all([first, second])).flat()).toMatchObject([
      { family: 'Noto Sans JP', bold: false, italic: false, kind: 'coverage' },
    ]);
    expect(source.resolveScriptFallback.mock.calls).toEqual(traits.map(([bold, italic]) => ['cjk-jp', bold, italic]));
    expect(install).toHaveBeenCalledTimes(1);
    expect(register).toHaveBeenCalledTimes(1);
    expect(settled).toHaveBeenCalledTimes(1);
    expect(fonts.isLoading()).toBe(false);
    expect(fonts.hasPending()).toBe(false);
    expect(await fonts.load()).toEqual([]);
  } finally { fonts.dispose(); }
});

test('a family requested as primary and as script coverage installs its face once', async () => {
  const japanese = await notoSansJP();
  const source = provider((family) => family === 'Noto Sans JP' ? async () => japanese : undefined,
    (script) => script === 'cjk-jp' ? async () => japanese : undefined);
  const register = mock((_face: DefaultFontFace) => {});
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces);
  const fonts = createCommandDefaultFonts([], () => [{ family: 'Noto Sans JP', bold: false, italic: false, scripts: ['cjk-jp'] }],
    register, () => {}, install);
  try {
    expect(await fonts.load()).toMatchObject([
      { family: 'Noto Sans JP', bold: false, italic: false },
      { family: 'Noto Sans JP', bold: false, italic: false, kind: 'coverage' },
    ]);
    expect(source.resolve).toHaveBeenCalledTimes(1);
    expect(source.resolveScriptFallback).toHaveBeenCalledTimes(1);
    expect(install).toHaveBeenCalledTimes(1);
    expect(register).toHaveBeenCalledTimes(2);
  } finally { fonts.dispose(); }
});

test('missing CJK packs never install Latin substitute bytes as coverage', async () => {
  const latin = new Uint8Array(await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))).buffer;
  const source = provider((family) => family === 'Calibri' ? async () => latin : undefined, () => async () => latin);
  const register = mock((_face: DefaultFontFace) => {});
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces);
  const fonts = createCommandDefaultFonts([], () => [
    { ...regular, scripts: ['cjk-jp'] }, { ...regular, family: 'BIZ UD明朝', scripts: ['cjk-jp'] },
  ], register, () => {}, install);
  try {
    expect(await fonts.load()).toMatchObject([{ family: 'Calibri' }]);
    expect(source.resolveScriptFallback.mock.calls).toEqual([['cjk-jp', false, false]]);
    expect(install).toHaveBeenCalledTimes(1);
    expect(register).toHaveBeenCalledTimes(1);
    expect(register.mock.calls[0][0]).not.toHaveProperty('kind', 'coverage');
  } finally { fonts.dispose(); }
});

test('browser refusal prevents coverage registration and settlement', async () => {
  const arabic = new Uint8Array(await readFile(new URL('../../../fonts/assets/NotoSansArabic-Regular.ttf', import.meta.url))).buffer;
  provider(() => undefined, () => async () => arabic);
  const register = mock(() => {});
  const settled = mock(() => {});
  const beforeRegister = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [{ ...regular, scripts: ['arabic'] }],
    register, settled, async () => [], beforeRegister);
  try {
    expect(await fonts.load()).toEqual([]);
    expect(register).not.toHaveBeenCalled();
    expect(settled).not.toHaveBeenCalled();
    expect(beforeRegister).not.toHaveBeenCalled();
  } finally { fonts.dispose(); }
});

test('with sliced faces off, a sliced loader registers the whole face and its exact bytes', async () => {
  const whole = new Uint8Array([9, 8, 7]).buffer;
  const open = spyOn(progressive, 'openSlicedFace');
  restorers.push(() => open.mockRestore());
  const load = Object.assign(mock(async () => whole), { slices: { ref: { format: 1, index: 'x', indexSha256: '' } } });
  provider(() => load as never);
  const registered: DefaultFontFace[] = [];
  const fonts = createCommandDefaultFonts([], () => [regular], (face) => { registered.push(face); }, () => {});
  try {
    await fonts.load();
    expect(open).not.toHaveBeenCalled();
    expect(load).toHaveBeenCalledTimes(1);
    expect(registered.map((face) => [...new Uint8Array(face.bytes as ArrayBuffer)])).toEqual([[9, 8, 7]]);
  } finally { fonts.dispose(); }
});

test('with sliced faces off, coverage prefetch fetches whole faces', async () => {
  const open = spyOn(progressive, 'openSlicedFace');
  restorers.push(() => open.mockRestore());
  const load = Object.assign(mock(async () => bytes), { slices: { ref: { format: 1, index: 'x', indexSha256: '' } } });
  provider(() => undefined, () => load as never);
  prefetchCoverageFaces(['cjk-jp'], Promise.resolve(), () => commandFontProvider.load());
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(load).toHaveBeenCalledTimes(1);
  expect(open).not.toHaveBeenCalled();
});

test('a host-supplied face under a coverage family is never sliced over', async () => {
  const skeleton = await notoSansJP();
  const load = Object.assign(mock(async () => skeleton), { slices: { ref: { format: 1, index: 'x', indexSha256: '' } } });
  provider(() => undefined, (script) => script === 'cjk-jp' ? load as never : undefined);
  const open = spyOn(progressive, 'openSlicedFace').mockResolvedValue({ revision: 1, bytes: () => skeleton,
    ensureText: async () => 1, prefetch() {} } as never);
  restorers.push(() => open.mockRestore());
  const browser = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { userAgent: 'Chrome/149.0' } });
  restorers.push(() => {
    if (browser) Object.defineProperty(globalThis, 'navigator', browser);
    else Reflect.deleteProperty(globalThis, 'navigator');
  });
  const sliced = createSlicedFontFaces();
  const fonts = createCommandDefaultFonts([{ family: 'Noto Sans JP' }], () => [{ ...regular, family: 'BIZ UD明朝', scripts: ['cjk-jp'] }],
    () => {}, () => {}, undefined, undefined, undefined, sliced);
  try {
    await fonts.load();
    expect(open).toHaveBeenCalled();
    expect(sliced.families()).toEqual([]);
  } finally { fonts.dispose(); sliced.dispose(); }
});

test('regular sliced with a whole substitute for bold: the family is sliced, and bold stays out of it', async () => {
  const skeleton = await notoSansJP();
  const sliceable = Object.assign(mock(async () => skeleton), { slices: { ref: { format: 1, index: 'x', indexSha256: '' } } });
  provider((_family, bold) => (bold ? async () => bytes : sliceable) as never);
  const open = spyOn(progressive, 'openSlicedFace').mockResolvedValue({ revision: 1, bytes: () => skeleton,
    ensureText: async () => 1, prefetch() {} } as never);
  restorers.push(() => open.mockRestore());
  const browser = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { userAgent: 'Chrome/149.0' } });
  restorers.push(() => {
    if (browser) Object.defineProperty(globalThis, 'navigator', browser);
    else Reflect.deleteProperty(globalThis, 'navigator');
  });
  const sliced = createSlicedFontFaces();
  const fonts = createCommandDefaultFonts([], () => [regular, { ...regular, bold: true }], () => {}, () => {},
    undefined, undefined, undefined, sliced);
  try {
    await fonts.load();
    expect(sliced.families()).toEqual(['Calibri']);
    expect(sliced.plainStyles()).toEqual(['calibri|1|0']);
    expect(sliced.coverage.covers!('Calibri', true, false)).toBe(false);
    expect(sliced.coverage.covers!('Calibri', false, false)).toBe(true);
  } finally { fonts.dispose(); sliced.dispose(); }
});

test('a worker provider applies its session CDN snapshot before creating loaders', async () => {
  const create = spyOn(fonts, 'createFontProvider');
  restorers.push(() => create.mockRestore());
  const configured = { baseUrl: 'https://worker-provider.example/', timeoutMs: 4321 };
  try {
    await commandFontProvider.load(configured);
    expect(fonts.fontCdnConfiguration()).toEqual(configured);
    await commandFontProvider.load({ baseUrl: false, timeoutMs: 30_000 });
    expect(fonts.fontCdnConfiguration().baseUrl).toBe(false);
    expect(create).toHaveBeenCalledTimes(2);
  } finally { fonts.configureFontCdn({ baseUrl: false }); }
});

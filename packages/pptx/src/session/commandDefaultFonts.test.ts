import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';
import { commandFontProvider, createCommandDefaultFonts } from './commandDefaultFonts';

const restorers: (() => void)[] = [];
const regular = { family: 'Calibri', bold: false, italic: false };
const bytes = new Uint8Array([1, 2, 3]).buffer;

function provider(
  load: (family: string, bold: boolean, italic: boolean) => (() => Promise<ArrayBuffer>) | undefined,
  script: (script: string, bold: boolean, italic: boolean) => (() => Promise<ArrayBuffer>) | undefined = () => undefined,
) {
  const resolve = mock(load);
  const resolveFamily = mock(() => undefined);
  const resolveScriptFallback = mock(script);
  const source = { resolve, resolveFamily, resolveScriptFallback,
    resolveLastResort: () => { throw new Error('Unrelated substitution'); } };
  const imported = spyOn(commandFontProvider, 'load').mockResolvedValue(source);
  restorers.push(() => imported.mockRestore());
  return { imported, resolve, resolveFamily, resolveScriptFallback };
}

afterEach(() => { for (const restore of restorers.splice(0).reverse()) restore(); });

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

test('resolves requested traits through manager mappings and skips unknown families', async () => {
  const source = provider((family) => family === 'Calibri' ? async () => bytes : undefined);
  const register = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [{ ...regular, bold: true, italic: true },
    { ...regular, family: 'Unknown Family' }], register, () => {});
  try {
    const faces = await fonts.load();
    expect(source.resolve.mock.calls).toEqual([['Calibri', true, true], ['Unknown Family', false, false]]);
    expect(source.resolveFamily).toHaveBeenCalledWith('Unknown Family', false, false);
    expect(faces).toHaveLength(1);
    expect(faces[0]).toMatchObject({ family: 'Calibri', bold: true, italic: true });
    structuredClone(faces, { transfer: faces.map((face) => face.bytes) });
    expect(bytes.byteLength).toBe(3);
  } finally { fonts.dispose(); }
});

test('lighter members of bundled families load their regular face; other unknown names stay unregistered', async () => {
  const resolve = mock(() => undefined);
  const resolveLightVariant = mock((family: string, _bold: boolean, _italic: boolean) =>
    family === 'Calibri Light' ? async () => bytes : undefined);
  const resolveLastResort = mock(() => { throw new Error('Unrelated substitution'); });
  const imported = spyOn(commandFontProvider, 'load').mockResolvedValue({ resolve, resolveFamily: () => undefined,
    resolveScriptFallback: () => undefined, resolveLastResort, resolveLightVariant });
  restorers.push(() => imported.mockRestore());
  const register = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [{ ...regular, family: 'Calibri Light', bold: true },
    { ...regular, family: 'BIZ UD明朝 Medium' }], register, () => {});
  try {
    const faces = await fonts.load();
    expect(resolveLightVariant.mock.calls).toEqual([['Calibri Light', true, false], ['BIZ UD明朝 Medium', false, false]]);
    expect(resolveLastResort).not.toHaveBeenCalled();
    expect(faces.map((face) => [face.family, face.bold])).toEqual([['Calibri Light', true]]);
    expect(register).toHaveBeenCalledTimes(1);
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
  const japanese = new Uint8Array(await readFile(new URL('../../../fonts-cjk/assets/NotoSansJP-Regular.otf', import.meta.url))).buffer;
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
  const japanese = new Uint8Array(await readFile(new URL('../../../fonts-cjk/assets/NotoSansJP-Regular.otf', import.meta.url))).buffer;
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

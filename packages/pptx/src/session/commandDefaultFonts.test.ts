import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import { sleep } from 'bun';
import { readFile } from 'node:fs/promises';
import * as fonts from '@betteroffice/fonts';
import { fontAssetResponse, fontRuntimeWithoutPacks, withBrowserFontGlobals } from '../../../fonts/src/fontAssets.test-support';
import { spyOnFontFetch } from '../../../fonts/src/fontFetch.test-support';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../../shared/office-session/fontCdnRuntime';
import * as progressive from '@betteroffice/fonts/progressive';
import { configureFontCdn as configureTestFontCdn } from '../../../fonts/src/fontCdn';
import type { FontCdnDiagnostic, FontCdnReporter } from '../../../../shared/fontCdnDiagnostics';
import type { DefaultFontFace, DefaultFontLanding } from '../../../../shared/pptxDefaultFonts';
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
  script: (script: string, bold: boolean, italic: boolean, family?: string) => (() => Promise<ArrayBuffer>) | undefined = () => undefined,
) {
  const resolve = mock(load);
  const resolveFamily = mock(() => undefined);
  const resolveScriptFallback = mock(script);
  const resolveLastResort = mock((_family: string, _bold: boolean, _italic: boolean, _office?: string) =>
    async (): Promise<ArrayBuffer> => { throw new Error('No substitute'); });
  const source = { resolve, resolveFamily, resolveScriptFallback, resolveLastResort };
  const imported = spyOn(commandFontProvider, 'load').mockImplementation(async () => ({ ...source }));
  restorers.push(() => imported.mockRestore());
  return { imported, resolve, resolveFamily, resolveScriptFallback, resolveLastResort };
}

function deadline() {
  const timeout = globalThis.setTimeout;
  const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: TimerHandler, delay?: number, ...args: unknown[]) =>
    timeout(callback, delay === 5000 ? 10 : delay, ...args)) as typeof setTimeout);
  restorers.push(() => timer.mockRestore());
}

function retryClock() {
  let now = Date.now();
  const clock = spyOn(Date, 'now').mockImplementation(() => now);
  restorers.push(() => clock.mockRestore());
  return (elapsed: number) => { now += elapsed; };
}

afterEach(() => { for (const restore of restorers.splice(0).reverse()) restore(); });

function packDiagnostics(): FontCdnDiagnostic[] {
  const diagnostics: FontCdnDiagnostic[] = [];
  const dispatch = Object.getOwnPropertyDescriptor(globalThis, 'dispatchEvent');
  Object.defineProperty(globalThis, 'dispatchEvent', {
    configurable: true,
    value(event: CustomEvent<FontCdnDiagnostic>) {
      if (event.type === 'betteroffice:font-pack-diagnostic') diagnostics.push(event.detail);
      return true;
    },
  });
  restorers.push(() => {
    if (dispatch) Object.defineProperty(globalThis, 'dispatchEvent', dispatch);
    else Reflect.deleteProperty(globalThis, 'dispatchEvent');
  });
  return diagnostics;
}

function implicitFontFailure(pack: FontCdnDiagnostic['pack'], cause: Error) {
  return Object.assign(new Error(cause.message), {
    cause, diagnostic: { pack, reason: 'cdn-unavailable' as const, detail: cause.message }, retryAt: Date.now() + 60_000,
  });
}

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

test('failed and invalid faces stay unregistered while timed out faces register late', async () => {
  let finish!: (value: ArrayBuffer) => void;
  const slow = new Promise<ArrayBuffer>((resolve) => { finish = resolve; });
  const slowLoad = mock(() => slow);
  provider((family) => family === 'Slow' ? slowLoad : family === 'Failure'
    ? async () => { throw new Error('Unavailable'); } : async () => bytes);
  deadline();
  const register = mock((face: { family: string }) => { if (face.family === 'Invalid') throw new Error('Invalid font'); });
  const settled = mock(() => {});
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces);
  const fonts = createCommandDefaultFonts([], () => ['Slow', 'Failure', 'Invalid', 'Calibri']
    .map((family) => ({ ...regular, family })), register, settled, install);
  try {
    expect(await fonts.load()).toMatchObject([{ family: 'Calibri' }]);
    expect(settled).toHaveBeenCalledTimes(1);
    expect(register.mock.calls.map(([face]) => face.family)).toEqual(['Invalid', 'Calibri']);
    expect(await fonts.load()).toEqual([]);
    finish(bytes);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(register.mock.calls.map(([face]) => face.family)).toEqual(['Invalid', 'Calibri', 'Slow']);
    expect(settled).toHaveBeenCalledTimes(2);
    expect(settled).toHaveBeenLastCalledWith([{ ...regular, family: 'Slow' }]);
    expect(install).toHaveBeenLastCalledWith([{ ...regular, family: 'Slow', bytes }]);
    expect(slowLoad).toHaveBeenCalledTimes(1);
    expect(await fonts.load()).toEqual([]);
    expect(settled).toHaveBeenCalledTimes(2);
  } finally { fonts.dispose(); }
});

test('a transient failure retries after backoff and registers the second load', async () => {
  const advance = retryClock();
  const load = mock(async () => bytes).mockRejectedValueOnce(new Error('Unavailable'));
  provider(() => load);
  const register = mock(() => {});
  const settled = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [regular], register, settled);
  try {
    expect(await fonts.load()).toEqual([]);
    expect(fonts.hasPending()).toBe(false);
    expect(await fonts.load()).toEqual([]);
    expect(load).toHaveBeenCalledTimes(1);
    advance(1999);
    expect(fonts.hasPending()).toBe(false);
    expect(await fonts.load()).toEqual([]);
    expect(load).toHaveBeenCalledTimes(1);
    advance(1);
    expect(fonts.hasPending()).toBe(true);
    expect(await fonts.load()).toEqual([regular]);
    expect(load).toHaveBeenCalledTimes(2);
    expect(register).toHaveBeenCalledTimes(1);
    expect(settled).toHaveBeenCalledTimes(1);
    expect(settled).toHaveBeenCalledWith([regular]);
  } finally { fonts.dispose(); }
});

test.each(['primary', 'last-resort'] as const)('a transient CDN failure retries the %s after the cooldown', async (kind) => {
  const advance = retryClock();
  const diagnostics = packDiagnostics();
  const load = mock(async () => bytes).mockRejectedValueOnce(implicitFontFailure('latin', new Error('Offline')));
  const primary = kind === 'primary' ? regular : { ...regular, family: 'Simplified Arabic', scripts: ['arabic'] };
  const source = provider(() => kind === 'primary' ? load : async () => bytes);
  if (kind === 'last-resort') source.resolveLastResort.mockReturnValue(load);
  const register = mock((_face: DefaultFontFace) => {});
  const manager = createCommandDefaultFonts([], () => [primary], register, () => {});
  try {
    const { family, bold, italic } = primary;
    expect(await manager.load()).toEqual(kind === 'primary' ? [] : [{ family, bold, italic }]);
    advance(59_999);
    expect(manager.hasPending()).toBe(false);
    expect(await manager.load()).toEqual([]);
    expect(load).toHaveBeenCalledTimes(1);
    advance(1);
    expect(manager.hasPending()).toBe(true);
    const landing: DefaultFontLanding = kind === 'primary' ? regular : { ...regular, family: 'Carlito', kind: 'coverage' };
    expect(await manager.load()).toEqual([landing]);
    expect(load).toHaveBeenCalledTimes(2);
    expect(register).toHaveBeenLastCalledWith({ ...landing, bytes });
    expect(manager.hasPending()).toBe(false);
    expect(diagnostics).toEqual([expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' })]);
  } finally { manager.dispose(); }
});

test.each([403, 404, 410])('a Naskh HTTP %s miss uses the available last resort as the primary', async (status) => {
  const advance = retryClock();
  const diagnostics = packDiagnostics();
  const naskh = mock(async () => bytes).mockRejectedValueOnce(
    implicitFontFailure('latin', Object.assign(new Error(`HTTP ${status}`), { status })));
  const source = provider(() => undefined, () => naskh);
  source.resolveLastResort.mockReturnValue(async () => bytes);
  const primary = { ...regular, family: 'Simplified Arabic' };
  const register = mock((_face: DefaultFontFace) => {});
  const manager = createCommandDefaultFonts([], () => [{ ...primary, scripts: ['arabic'] }], register, () => {});
  try {
    expect(await manager.load()).toContainEqual(primary);
    expect(register).toHaveBeenCalledWith({ ...primary, bytes });
    advance(60_000);
    expect(manager.hasPending()).toBe(false);
    expect(await manager.load()).toEqual([]);
    expect(naskh).toHaveBeenCalledTimes(1);
    expect(diagnostics).toEqual([expect.objectContaining({ pack: 'latin', reason: 'cdn-unavailable' })]);
  } finally { manager.dispose(); }
});

test.each([403, 404, 410])('an HTTP %s retry keeps the installed substitute without reinstalling it', async (status) => {
  const advance = retryClock();
  const diagnostics = packDiagnostics();
  let cause = new Error('Offline');
  const load = mock(async () => bytes.slice(0));
  const create = spyOn(fonts, 'createFontProvider').mockImplementation((_options, runtime) => ({
    resolve: () => async () => {
      const hooks = runtime as (NonNullable<typeof runtime> & {
        reportCdnFailure?: FontCdnReporter;
        onSubstitute?: (bytes: ArrayBuffer, error: unknown) => void;
      }) | undefined;
      const substitute = await load();
      const error = implicitFontFailure('intl', cause);
      hooks?.reportCdnFailure?.(error);
      hooks?.onSubstitute?.(substitute, error);
      return substitute;
    },
    resolveFamily: () => undefined,
    resolveScriptFallback: () => undefined,
    resolveLastResort: () => async () => bytes,
  }));
  restorers.push(() => create.mockRestore());
  const primary = { ...regular, family: 'Noto Sans Devanagari' };
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces);
  const register = mock((_face: DefaultFontFace) => {});
  const manager = createCommandDefaultFonts([], () => [primary], register, () => {}, install);
  try {
    expect(await manager.load()).toEqual([primary]);
    cause = Object.assign(new Error(`HTTP ${status}`), { status });
    advance(60_000);
    const additions = await manager.load();
    expect(load).toHaveBeenCalledTimes(2);
    expect(install).toHaveBeenCalledTimes(1);
    expect(register).toHaveBeenCalledTimes(1);
    expect(additions).toEqual([]);
    advance(60_000);
    expect(manager.hasPending()).toBe(false);
    expect(diagnostics).toEqual([expect.objectContaining({ pack: 'intl', reason: 'cdn-unavailable' })]);
  } finally { manager.dispose(); }
});

test.each([false, true])('an absent intl pack settles without fetching a substitute (CDN blocked: %s)', async (blocked) => {
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals(false);
  const cdn = { baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit: true } as const;
  const advance = retryClock();
  const requests: string[] = [];
  const fetch = spyOnFontFetch(async (input) => {
    requests.push(String(input));
    if (String(input).includes('/fonts-intl@')) return new Response(null, { status: 403 });
    return fontAssetResponse(input);
  });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const runtime = fontRuntimeWithoutPacks(new URL('https://script-fallback.example/'));
  const create = spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
  const register = mock(() => {});
  const manager = createCommandDefaultFonts([regular], () => [{ ...regular, scripts: ['thai'] }], register, () => {},
    undefined, undefined, () => commandFontProvider.load(blocked ? cdn : { baseUrl: false, timeoutMs: 10_000 }));
  try {
    expect(await manager.load()).toEqual([]);
    advance(60_001);
    expect(manager.hasPending()).toBe(false);
    expect(await manager.load()).toEqual([]);
    expect(register).not.toHaveBeenCalled();
    expect(requests).toEqual(blocked ? [expect.stringContaining('/fonts-intl@')] : []);
    expect(warn).toHaveBeenCalledTimes(blocked ? 1 : 0);
  } finally {
    manager.dispose();
    create.mockRestore();
    fetch.mockRestore();
    warn.mockRestore();
    restore();
    setFontCdnConfiguration(previous);
  }
});

test('primary and coverage misses emit one diagnostic per pack in each session', async () => {
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals(false);
  const cdn = { baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit: true } as const;
  const runtime = fontRuntimeWithoutPacks(new URL('https://primary-coverage.example/'));
  const create = spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
  const fetch = spyOnFontFetch((input) => String(input).includes('/fonts-intl@')
    ? Promise.resolve(new Response(null, { status: 403 })) : fontAssetResponse(input));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const diagnostics = packDiagnostics();
  const primary = { ...regular, family: 'Noto Sans Devanagari' };
  const diagnostic = expect.objectContaining({ pack: 'intl', reason: 'cdn-unavailable' });
  try {
    for (let session = 1; session <= 2; session += 1) {
      const manager = createCommandDefaultFonts([regular], () => [primary, { ...regular, scripts: ['thai'] }],
        () => {}, () => {}, undefined, undefined, () => commandFontProvider.load(cdn));
      try {
        expect(await manager.load()).toEqual([primary]);
        expect(diagnostics).toEqual(Array(session).fill(diagnostic));
      } finally { manager.dispose(); }
    }
    expect(fetch.mock.calls.filter(([url]) => String(url).includes('/fonts-intl@'))).toHaveLength(1);
  } finally {
    create.mockRestore();
    fetch.mockRestore();
    warning.mockRestore();
    restore();
    setFontCdnConfiguration(previous);
  }
});

test.each([403, 404, 410])('an implicit HTTP %s miss settles the preferred Arabic primary', async (status) => {
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals(false);
  const cdn = { baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit: true } as const;
  const advance = retryClock();
  const fetch = spyOnFontFetch(async () => new Response(null, { status }));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const create = spyOn(fonts, 'createFontProvider').mockImplementation(fontRuntimeWithoutPacks().createFontProvider);
  const manager = createCommandDefaultFonts([], () => [{ ...regular, family: 'Simplified Arabic', scripts: ['arabic'] }],
    () => {}, () => {}, undefined, undefined, () => commandFontProvider.load(cdn));
  try {
    expect(await manager.load()).toEqual([]);
    const requests = fetch.mock.calls.length;
    expect(requests).toBeGreaterThan(0);
    advance(60_001);
    expect(manager.hasPending()).toBe(false);
    expect(await manager.load()).toEqual([]);
    expect(fetch).toHaveBeenCalledTimes(requests);
  } finally {
    manager.dispose();
    create.mockRestore();
    fetch.mockRestore();
    warning.mockRestore();
    restore();
    setFontCdnConfiguration(previous);
  }
});

test('a primary installed with substitute bytes installs the requested face after the CDN recovers', async () => {
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals(false);
  const cdn = { baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit: true } as const;
  const advance = retryClock();
  let online = false;
  const fetch = spyOnFontFetch(async (input) => {
    if (!online && String(input).includes('/fonts-intl@')) throw new TypeError('Offline');
    return fontAssetResponse(input);
  });
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const create = spyOn(fonts, 'createFontProvider').mockImplementation(fontRuntimeWithoutPacks().createFontProvider);
  const primary = { ...regular, family: 'Noto Sans Devanagari' };
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces);
  const register = mock((_face: DefaultFontFace) => {});
  const manager = createCommandDefaultFonts([], () => [primary], register, () => {}, install,
    undefined, () => commandFontProvider.load(cdn));
  try {
    expect(await manager.load()).toEqual([primary]);
    expect(install).toHaveBeenCalledTimes(1);
    advance(59_999);
    expect(manager.hasPending()).toBe(false);
    advance(1);
    expect(await manager.load()).toEqual([]);
    expect(install).toHaveBeenCalledTimes(1);
    online = true;
    advance(60_000);
    expect(await manager.load()).toEqual([primary]);
    expect(install).toHaveBeenCalledTimes(2);
    expect(register).toHaveBeenCalledTimes(2);
    expect(install.mock.calls[1][0][0].bytes).not.toEqual(install.mock.calls[0][0][0].bytes);
    expect(manager.hasPending()).toBe(false);
  } finally {
    manager.dispose();
    create.mockRestore();
    fetch.mockRestore();
    warning.mockRestore();
    restore();
    setFontCdnConfiguration(previous);
  }
});

test('an implicit HTTP 403 settles other faces in the same pack after the CDN cooldown', async () => {
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals(false);
  const cdn = { baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit: true } as const;
  const advance = retryClock();
  const fetch = spyOnFontFetch(async () => new Response(null, { status: 403 }));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const diagnostics = packDiagnostics();
  const runtime = fontRuntimeWithoutPacks();
  const decoder: fonts.FontDecoder = (data, options) => decodeFontFace(data, options);
  const create = spyOn(fonts, 'createFontProvider').mockImplementation((options, loadOptions) =>
    runtime.createFontProvider(options, { ...loadOptions, decoder }));
  const collect = (slide?: string) => [{ ...regular, scripts: [slide === 'next' ? 'devanagari' : 'thai'] }];
  const manager = createCommandDefaultFonts([regular], collect, () => {}, () => {}, undefined, undefined,
    () => commandFontProvider.load(cdn));
  const nextSession = createCommandDefaultFonts([regular], collect, () => {}, () => {}, undefined, undefined,
    () => commandFontProvider.load(cdn));
  try {
    expect(await manager.load()).toEqual([]);
    expect(fetch.mock.calls.map(([url]) => String(url))).toEqual([expect.stringContaining('/NotoSansThai-Regular.ttf.zst')]);
    advance(60_001);
    expect(await manager.load('next')).toEqual([]);
    expect(manager.hasPending('next')).toBe(false);
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(warning).toHaveBeenCalledTimes(1);
    expect(diagnostics).toHaveLength(1);

    expect(await nextSession.load('next')).toEqual([]);
    expect(fetch.mock.calls.map(([url]) => String(url))).toEqual([
      expect.stringContaining('/NotoSansThai-Regular.ttf.zst'), expect.stringContaining('/NotoSansDevanagari-Regular.ttf.zst'),
    ]);
    expect(warning).toHaveBeenCalledTimes(2);
    expect(diagnostics).toHaveLength(2);
  } finally {
    manager.dispose();
    nextSession.dispose();
    create.mockRestore();
    fetch.mockRestore();
    warning.mockRestore();
    restore();
    setFontCdnConfiguration(previous);
  }
});

const slicedMisses = [403, 404, 410].flatMap<[number, boolean, boolean]>((status) => [
  [status, false, true], [status, false, false], [status, true, false],
]);

test.each(slicedMisses)('a sliced HTTP %s miss tries the whole face when allowed (implicit: %s, whole available: %s)', async (status, implicit, wholeAvailable) => {
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals(false);
  const runtime = fontRuntimeWithoutPacks();
  const create = spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
  const whole = fonts.resolveScriptFallbackFace('cjk-jp', false, false)!.file;
  const fetch = spyOnFontFetch(async (input) => wholeAvailable && String(input).endsWith(`/${whole}`)
    ? fontAssetResponse(input) : new Response(null, { status }));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const diagnostics = packDiagnostics();
  const sliced = createSlicedFontFaces();
  const navigator = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { userAgent: 'Chrome/149.0' } });
  const advance = retryClock();
  const wanted = { ...regular, scripts: ['cjk-jp'] };
  const cdn = implicit ? { baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit }
    : { baseUrl: `https://sliced-miss.example/${status}/${wholeAvailable ? 'whole' : 'missing'}/`, timeoutMs: 10_000 };
  let source: ReturnType<typeof commandFontProvider.load> | undefined;
  const shared = () => (source ??= commandFontProvider.load(cdn));
  const manager = createCommandDefaultFonts([regular], () => [wanted], () => {}, () => {},
    undefined, undefined, shared, sliced);
  try {
    prefetchCoverageFaces(['cjk-jp'], Promise.resolve(), shared, true);
    await sleep(0);
    expect(await manager.load()).toEqual(wholeAvailable ? [{ ...regular, family: 'Noto Sans JP', kind: 'coverage' }] : []);
    advance(60_001);
    expect(manager.hasPending()).toBe(false);
    expect(await manager.load()).toEqual([]);
    expect(fetch.mock.calls.map(([url]) => String(url))).toEqual([
      expect.stringMatching(/\/slices\.json$/), ...(implicit ? [] : [expect.stringContaining(`/${whole}`)]),
    ]);
    expect(diagnostics).toEqual(implicit ? [expect.objectContaining({ pack: 'cjk', reason: 'cdn-unavailable' })] : []);
  } finally {
    manager.dispose();
    sliced.dispose();
    if (navigator) Object.defineProperty(globalThis, 'navigator', navigator);
    else Reflect.deleteProperty(globalThis, 'navigator');
    create.mockRestore();
    fetch.mockRestore();
    warning.mockRestore();
    restore();
    setFontCdnConfiguration(previous);
  }
});

test('coverage prefetch and slide switches share a CDN failure and retry in the same session', async () => {
  const previous = getFontCdnConfiguration();
  configureTestFontCdn({ baseUrl: false });
  const restore = withBrowserFontGlobals(false);
  const advance = retryClock();
  const requests: string[] = [];
  let online = false;
  const fetch = spyOnFontFetch(async (input) => {
    if (String(input).includes('cdn.jsdelivr.net')) {
      requests.push(String(input));
      if (!online) throw new TypeError('Offline');
    }
    return fontAssetResponse(input);
  });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const current = { ...regular, scripts: ['cjk-jp'] };
  const next = { ...regular, scripts: ['cjk-kr'] };
  const register = mock(() => {});
  setFontCdnConfiguration({ baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit: true });
  const source = fontRuntimeWithoutPacks().createFontProvider(undefined, { strict: true, retain: false });
  const shared = async () => source;
  const manager = createCommandDefaultFonts([regular], (slide) => [slide === 'next' ? next : current],
    register, () => {}, undefined, undefined, shared);
  try {
    prefetchCoverageFaces(['cjk-jp'], Promise.resolve(), shared);
    await sleep(0);
    expect(requests).toHaveLength(1);
    expect(await manager.load('opening')).toEqual([]);
    expect(await manager.load('next')).toEqual([]);
    expect(requests).toHaveLength(1);
    expect(register).not.toHaveBeenCalled();
    expect(manager.hasPending('opening')).toBe(false);
    expect(manager.hasPending('next')).toBe(false);

    online = true;
    advance(60_000);
    expect(await manager.load('opening')).toHaveLength(1);
    expect(requests).toHaveLength(2);
    expect(register).toHaveBeenCalledTimes(1);
  } finally {
    manager.dispose();
    fetch.mockRestore();
    warn.mockRestore();
    configureTestFontCdn({ baseUrl: false });
    restore();
    setFontCdnConfiguration(previous);
  }
});

test('rejected faces back off for two and eight seconds and stop after three attempts', async () => {
  const advance = retryClock();
  const load = mock(async (): Promise<ArrayBuffer> => { throw new Error('Unavailable'); });
  provider(() => load);
  const register = mock(() => {});
  const settled = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [regular], register, settled);
  try {
    expect(await fonts.load()).toEqual([]);
    for (const [delay, attempts] of [[2000, 1], [8000, 2]]) {
      advance(delay - 1);
      expect(fonts.hasPending()).toBe(false);
      for (let call = 0; call < 20; call++) expect(await fonts.load()).toEqual([]);
      expect(load).toHaveBeenCalledTimes(attempts);
      advance(1);
      expect(fonts.hasPending()).toBe(true);
      expect(await fonts.load()).toEqual([]);
      expect(load).toHaveBeenCalledTimes(attempts + 1);
    }
    for (let call = 0; call < 20; call++) {
      advance(60_000);
      expect(fonts.hasPending()).toBe(false);
      expect(await fonts.load()).toEqual([]);
    }
    expect(load).toHaveBeenCalledTimes(3);
    expect(register).not.toHaveBeenCalled();
    expect(settled).not.toHaveBeenCalled();
  } finally { fonts.dispose(); }
});

test('a face with no loader stays attempted without retries', async () => {
  const advance = retryClock();
  const source = provider(() => undefined);
  source.resolveLastResort.mockReturnValue(undefined as never);
  const register = mock(() => {});
  const settled = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [regular], register, settled);
  try {
    expect(await fonts.load()).toEqual([]);
    advance(60_000);
    expect(fonts.hasPending()).toBe(false);
    expect(await fonts.load()).toEqual([]);
    expect(source.resolve).toHaveBeenCalledTimes(1);
    expect(register).not.toHaveBeenCalled();
    expect(settled).not.toHaveBeenCalled();
  } finally { fonts.dispose(); }
});

test('a refused installation is retried after backoff', async () => {
  const advance = retryClock();
  const load = mock(async () => bytes);
  provider(() => load);
  const register = mock(() => {});
  const settled = mock(() => {});
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces).mockResolvedValueOnce([]);
  const fonts = createCommandDefaultFonts([], () => [regular], register, settled, install);
  try {
    expect(await fonts.load()).toEqual([]);
    expect(register).not.toHaveBeenCalled();
    advance(2000);
    expect(await fonts.load()).toEqual([regular]);
    expect(load).toHaveBeenCalledTimes(2);
    expect(install).toHaveBeenCalledTimes(2);
    expect(register).toHaveBeenCalledTimes(1);
    expect(settled).toHaveBeenCalledTimes(1);
    expect(settled).toHaveBeenCalledWith([regular]);
  } finally { fonts.dispose(); }
});

test('a face that never resolves releases first paint at the deadline without refetching', async () => {
  deadline();
  const load = mock(() => new Promise<ArrayBuffer>(() => {}));
  provider(() => load);
  const register = mock(() => {});
  const settled = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [regular], register, settled);
  try {
    expect(await fonts.load()).toEqual([]);
    expect(await fonts.load()).toEqual([]);
    expect(load).toHaveBeenCalledTimes(1);
    expect(register).not.toHaveBeenCalled();
    expect(settled).not.toHaveBeenCalled();
  } finally { fonts.dispose(); }
});

for (const action of ['supply', 'dispose'] as const) test(`${action} prevents registration after the deadline`, async () => {
  deadline();
  let finish!: (value: ArrayBuffer) => void;
  provider(() => () => new Promise<ArrayBuffer>((resolve) => { finish = resolve; }));
  const register = mock(() => {});
  const settled = mock(() => {});
  const fonts = createCommandDefaultFonts([], () => [regular], register, settled);
  try {
    expect(await fonts.load()).toEqual([]);
    if (action === 'supply') fonts.supply(' CALIBRI ');
    else fonts.dispose();
    finish(bytes);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(register).not.toHaveBeenCalled();
    expect(settled).not.toHaveBeenCalled();
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

test('Arabic requests keep family and traits and install the shared regular face only once', async () => {
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
    expect(source.resolveScriptFallback.mock.calls).toEqual([
      ['arabic', true, true, 'Calibri'], ['arabic', false, false, 'Calibri'],
      ['arabic', true, true, 'Unknown Family'], ['arabic', false, false, 'Unknown Family'],
    ]);
    expect(results.flat()).toMatchObject([{ family: 'Noto Sans Arabic', bold: false, italic: false, kind: 'coverage' }]);
    expect(install).toHaveBeenCalledTimes(1);
    expect(install.mock.calls[0][0][0].kind).toBe('coverage');
    expect(register).toHaveBeenCalledTimes(1);
    expect(settled).toHaveBeenCalledTimes(1);
    expect(fonts.hasPending()).toBe(false);
    expect(await fonts.load()).toEqual([]);
  } finally { fonts.dispose(); }
});

const preferredArabicStyles = ['Simplified Arabic', 'Traditional Arabic'].flatMap((family) =>
  [[true, false], [false, true], [true, true]].map(([bold, italic]): [string, boolean, boolean] => [family, bold, italic]));

test.each(preferredArabicStyles)('%s uses the requested browser style (bold: %s, italic: %s)', async (family, bold, italic) => {
  const naskh = new Uint8Array(await readFile(new URL('../../../fonts/assets/NotoNaskhArabic-Regular.ttf', import.meta.url))).buffer;
  const source = provider(() => undefined, () => async () => naskh);
  source.resolveLastResort.mockReturnValue(async () => bytes);
  const install = mock(async (faces: readonly DefaultFontFace[]) => faces);
  const primary = { family, bold, italic };
  const manager = createCommandDefaultFonts([], () => [{ ...primary, scripts: ['arabic'] }], () => {}, () => {}, install);
  try {
    expect(await manager.load()).toContainEqual({ ...primary, ownWeight: true });
    expect(install).toHaveBeenCalledWith([{ ...primary, bytes: naskh, ownWeight: true }]);
  } finally { manager.dispose(); }
});

test.each([false, true])('Simplified Arabic uses its last resort while the Naskh face is unavailable (implicit: %s)', async (implicit) => {
  const naskh = new Uint8Array(await readFile(new URL('../../../fonts/assets/NotoNaskhArabic-Regular.ttf', import.meta.url))).buffer;
  const advance = retryClock();
  const delay = implicit ? 60_000 : 2000;
  const outage = new Error('Offline');
  if (implicit) Object.assign(outage, {
    diagnostic: { pack: 'latin', reason: 'cdn-unavailable', detail: outage.message }, retryAt: Date.now() + delay,
  });
  const load = mock(async () => naskh).mockRejectedValueOnce(outage);
  const source = provider(() => undefined, () => load);
  source.resolveLastResort.mockReturnValue(async () => bytes);
  const primary = { ...regular, family: 'Simplified Arabic' };
  const registered: DefaultFontFace[] = [];
  const manager = createCommandDefaultFonts([], () => [{ ...primary, scripts: ['arabic'] }],
    (face) => registered.push(face), () => {});
  try {
    expect(await manager.load()).toContainEqual(primary);
    expect(registered.filter((face) => face.family === primary.family)).toEqual([{ ...primary, bytes }]);
    advance(delay - 1);
    expect(manager.hasPending()).toBe(false);
    expect(await manager.load()).toEqual([]);
    expect(load).toHaveBeenCalledTimes(1);
    advance(1);
    expect(manager.hasPending()).toBe(true);
    expect(await manager.load()).toEqual([primary]);
    expect(registered.filter((face) => face.family === primary.family)).toEqual([{ ...primary, bytes }, { ...primary, bytes: naskh }]);
    expect(load).toHaveBeenCalledTimes(2);
    expect(manager.hasPending()).toBe(false);
  } finally { manager.dispose(); }
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
    await commandFontProvider.load(undefined);
    expect(getFontCdnConfiguration()).toBeUndefined();
    expect(create).toHaveBeenCalledTimes(3);
  } finally { fonts.configureFontCdn({ baseUrl: false }); }
});

test('a PPTX worker font provider preserves the implicit CDN marker from session prefetch', async () => {
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals(false);
  const inherited = { baseUrl: 'https://cdn.jsdelivr.net/npm/', timeoutMs: 10_000, implicit: true } as const;
  try {
    await commandFontProvider.load(inherited);
    expect(fonts.fontCdnConfiguration()).toEqual(inherited);
  } finally { restore(); setFontCdnConfiguration(previous); }
});

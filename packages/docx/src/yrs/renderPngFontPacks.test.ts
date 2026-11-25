import { afterEach, beforeAll, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as fonts from '@betteroffice/fonts';
import { LATIN_FONT_ASSET_URLS } from '../../../fonts-latin/src';
import { BASE_LATIN_FONT_ASSET_URLS } from '../../../fonts/src/base-latin';
import { fontAssetResponse, fontRuntimeWithoutPacks, withBrowserFontGlobals } from '../../../fonts/src/fontAssets.test-support';
import { spyOnFontFetch } from '../../../fonts/src/fontFetch.test-support';
import { createFontRuntime } from '../../../fonts/src/fontRuntime';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import { configureDefaultFonts, resolveCompleteFontProvider, resolveDefaultFontProvider } from '../layout/measure';
import { getGoogleFontEquivalent } from '../utils/fontLoader';
import { decodeFontFace } from '../wasm/fontDecoder';
import { engineKind } from '../wasm/enginePng';
import { createYrsSession, type YrsSession } from './index';
import { ENGINES, preloadEngines, restoreBrowser, runsDocument, scriptDocument, stubBrowser, type TextRun } from './renderPngTestSupport';

const LATIN_ASSETS = new URL('../../../fonts-latin/assets/', import.meta.url);
const withheld = new Set<string>();
const asked = new Set<string>();
let gatedAssets: URL | undefined;
let gates = 0;

mock.module('@betteroffice/fonts-latin', () => ({
  LATIN_FONT_ASSET_URLS: new Proxy({} as Record<string, () => URL>, {
    get(_, file) {
      if (typeof file !== 'string' || !file.endsWith('.ttf.zst')) return undefined;
      asked.add(file);
      return withheld.has(file) ? undefined : () => new URL(file, gatedAssets ?? LATIN_ASSETS);
    },
  }),
}));

const originalFetch = globalThis.fetch;

/** Serves the Latin pack from an origin of its own, so no face this file decoded before is reused. */
function gateLatinFetches(serve: (file: string) => Promise<Response>): void {
  const assets = new URL(`https://pack-${++gates}.gated.test/`);
  gatedAssets = assets;
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(input instanceof Request ? input.url : String(input));
    return url.origin === assets.origin ? serve(url.pathname.slice(1)) : originalFetch(input, init);
  }) as typeof fetch;
}

const latinBytes = (file: string) => new Response(readFileSync(new URL(file, LATIN_ASSETS)));
const arabicFace = () => fonts.resolveScriptFallbackFace('arabic', false, false, 'Arial')!;
const latinAndArabic: TextRun[] = [{ family: 'Arial', text: 'Plain text' }, { family: 'Arial', text: 'مرحبا بالعالم', rtl: true }];
const gothicRun = { family: 'MS Gothic', text: 'Plain text', paragraphFamily: 'MS Gothic' };

function mockFontsWithLatinPack() {
  const runtime = createFontRuntime(async (pack) => ({
    cdn: getFontCdnConfiguration(),
    urls: pack === 'latin' ? LATIN_FONT_ASSET_URLS : pack === 'base' ? BASE_LATIN_FONT_ASSET_URLS : undefined,
  }));
  return spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
}

beforeAll(preloadEngines);
afterEach(() => {
  withheld.clear();
  asked.clear();
  gatedAssets = undefined;
  globalThis.fetch = originalFetch;
  restoreBrowser();
});

async function exported(runs: TextRun[], options: Parameters<YrsSession['renderPng']>[1] = {}): Promise<Uint8Array> {
  const session = await createYrsSession();
  try {
    session.openDocx(runsDocument(...runs), true);
    return await session.renderPng({ page: 0 }, options);
  } finally { session.destroy(); }
}

async function rejection(work: Promise<unknown>): Promise<{ name?: string; code?: string; message: string }> {
  const error = await work.then(() => undefined, (reason: unknown) => reason);
  expect(error).toBeInstanceOf(Error);
  return error as { name?: string; code?: string; message: string };
}

for (const { kind } of ENGINES) {
  describe.skipIf(engineKind !== kind)(`renderPng font packs on the ${kind} engine`, () => {
    test('DOCX PNG rejects an implicit Latin CDN failure after one request', async () => {
      const canvas = stubBrowser();
      const previous = getFontCdnConfiguration();
      const restore = withBrowserFontGlobals();
      const provider = spyOn(fonts, 'createFontProvider').mockImplementation(fontRuntimeWithoutPacks().createFontProvider);
      const family = 'Calibri';
      const face = fonts.resolveMetricCompatFace(family, false, false)!;
      const requests: string[] = [];
      globalThis.fetch = (async (input: RequestInfo | URL) => {
        if (String(input).includes('cdn.jsdelivr.net')) {
          requests.push(String(input));
          return new Response(null, { status: 404 });
        }
        return fontAssetResponse(input);
      }) as typeof fetch;
      try {
        setFontCdnConfiguration(undefined);
        const error = await rejection(exported([{ family, text: 'Plain text', paragraphFamily: family }]));
        expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable', message: expect.stringContaining('HTTP 404') });
        expect(requests).toEqual([expect.stringContaining('/fonts-latin@')]);
        expect(requests[0]).toEndWith(`/assets/${face.file}`);
        expect(canvas.encoded).toBe(0);
      } finally {
        provider.mockRestore();
        restore();
        setFontCdnConfiguration(previous);
      }
    });

    test('DOCX PNG draws as with an absent CJK pack after the editor family load hits a blocked default CDN', async () => {
      const session = await createYrsSession();
      const previous = getFontCdnConfiguration();
      const restore = withBrowserFontGlobals();
      const create = mockFontsWithLatinPack();
      const requests: string[] = [];
      const fetch = spyOnFontFetch(async (input) => {
        if (String(input).startsWith('https://cdn.jsdelivr.net/')) {
          requests.push(String(input));
          throw new TypeError('Offline');
        }
        return fontAssetResponse(input);
      });
      try {
        configureDefaultFonts({});
        session.openDocx(runsDocument(gothicRun), true);
        const absentPackDrawing = stubBrowser();
        const absentPackContours = spyOn(Path2D.prototype, 'moveTo');
        fonts.configureFontCdn({ baseUrl: false });
        const absentPack = await session.renderPng({ page: 0 });
        expect(absentPack.byteLength).toBeGreaterThan(0);
        expect(absentPack.byteLength).toBeLessThan(1_000_000);
        expect(requests).toEqual([]);

        setFontCdnConfiguration(undefined);
        const provider = await resolveDefaultFontProvider();
        const family = getGoogleFontEquivalent(gothicRun.family);
        const load = provider!.resolveFamily!(family, false, false)!;
        await expect(load()).rejects.toMatchObject({
          name: 'ImplicitFontCdnError', diagnostic: { pack: 'cjk', reason: 'cdn-unavailable' },
        });
        const { file } = fonts.resolveMetricCompatFace(gothicRun.family, false, false)!;
        expect(requests).toEqual([expect.stringContaining('/fonts-cjk@')]);
        expect(requests[0]).toEndWith(`/assets/${file}`);

        const blockedCdnDrawing = stubBrowser();
        const blockedCdnContours = spyOn(Path2D.prototype, 'moveTo');
        await expect(session.renderPng({ page: 0 })).resolves.toEqual(absentPack);
        expect(requests).toHaveLength(1);
        if (kind === 'wasm') {
          expect(absentPackContours).toHaveBeenCalled();
          expect(blockedCdnContours.mock.calls).toEqual(absentPackContours.mock.calls);
          expect(blockedCdnDrawing).toEqual(absentPackDrawing);
        }
      } finally {
        session.destroy();
        fetch.mockRestore();
        create.mockRestore();
        restore();
        setFontCdnConfiguration(previous);
        configureDefaultFonts({});
      }
    });

    test('DOCX PNG rejects a failed CJK family load from an explicitly configured CDN', async () => {
      stubBrowser();
      const previous = getFontCdnConfiguration();
      const restore = withBrowserFontGlobals();
      const create = mockFontsWithLatinPack();
      const baseUrl = 'https://docx-png-cjk.example/';
      const fetch = spyOnFontFetch(async (input) => String(input).startsWith(baseUrl)
        ? new Response(null, { status: 503 })
        : fontAssetResponse(input));
      try {
        fonts.configureFontCdn({ baseUrl });
        await expect(exported([gothicRun])).rejects.toMatchObject({
          name: 'PngExportError',
          code: 'font-unavailable',
          message: expect.stringContaining('Noto Sans JP for MS Gothic could not load'),
        });
        const { file } = fonts.resolveMetricCompatFace(gothicRun.family, false, false)!;
        expect(fetch.mock.calls.some(([input]) => {
          const url = String(input);
          return url.startsWith(`${baseUrl}cjk/`) && url.endsWith(`/${file}`);
        })).toBe(true);
      } finally {
        fetch.mockRestore();
        create.mockRestore();
        restore();
        setFontCdnConfiguration(previous);
      }
    });

    test('DOCX PNG draws with supplied font bytes offline and rejects missing families', async () => {
      const canvas = stubBrowser();
      const previous = getFontCdnConfiguration();
      const restore = withBrowserFontGlobals();
      const runtime = fontRuntimeWithoutPacks(new URL('https://docx-offline-export-base.example/'));
      const provider = spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
      const fetch = spyOnFontFetch(async (input) => {
        if (String(input).includes('cdn.jsdelivr.net')) throw new TypeError('Offline');
        return fontAssetResponse(input);
      });
      try {
        setFontCdnConfiguration(undefined);
        const bytes = new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'));
        const calibri = { family: 'Calibri', text: 'Total', paragraphFamily: 'Calibri' };
        const arial = { family: 'Arial', text: 'Balance', paragraphFamily: 'Arial' };
        const supplied = { fonts: [{ family: ' calibri ', bytes }] };
        await expect(exported([calibri], supplied)).resolves.toBeInstanceOf(Uint8Array);
        expect(fetch).not.toHaveBeenCalled();
        if (kind === 'wasm') expect(canvas.fills).toBeGreaterThan(0);

        await expect(exported([calibri])).rejects.toMatchObject({
          code: 'font-unavailable', message: expect.stringContaining('Carlito for Calibri could not load'),
        });
        await expect(exported([calibri, arial], supplied)).rejects.toMatchObject({
          code: 'font-unavailable', message: expect.stringContaining('Liberation Sans for Arial could not load'),
        });
      } finally {
        fetch.mockRestore();
        provider.mockRestore();
        restore();
        setFontCdnConfiguration(previous);
      }
    });

    test.skipIf(kind !== 'wasm')('DOCX PNG draws Han glyphs with supplied Latin-only Arial', async () => {
      const face = fonts.resolveMetricCompatFace('Arial', false, false)!;
      const bytes = new Uint8Array(await fonts.loadBundledFontBytes(face, undefined, { decoder: decodeFontFace, strict: true }));
      const run = { family: 'Arial', text: 'Ab 漢字 cd', paragraphFamily: 'Arial' };
      stubBrowser();
      const contours = spyOn(Path2D.prototype, 'moveTo');
      try {
        await exported([run]);
        expect(contours).toHaveBeenCalled();
        const expected = [...contours.mock.calls];
        contours.mockClear();
        await exported([run], { fonts: [{ family: 'Arial', bytes }] });
        expect(contours.mock.calls).toEqual(expected);
      } finally { contours.mockRestore(); }
    });

    test.each(['default', 'supplied'] as const)('DOCX PNG with absent optional packs and %s fonts', async (source) => {
      stubBrowser();
      const previous = getFontCdnConfiguration();
      const runtime = fontRuntimeWithoutPacks(new URL('https://docx-headless-base.example/'));
      const provider = spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
      const fetch = spyOnFontFetch(fontAssetResponse);
      try {
        setFontCdnConfiguration(undefined);
        const calibri = { family: 'Calibri', text: 'Plain text', paragraphFamily: 'Calibri' };
        if (source === 'supplied') {
          const bytes = new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'));
          const png = await exported([calibri], { fonts: [{ family: 'Calibri', bytes }] });
          expect(png).toBeInstanceOf(Uint8Array);
          expect(png.byteLength).toBeGreaterThan(0);
          expect(fetch).not.toHaveBeenCalled();
          return;
        }
        if (kind === 'native') {
          const error = await rejection(exported([calibri]));
          expect(error.name).toBe('PngExportError');
          expect(error.code).toBe('font-unavailable');
          expect(error.message).toContain('missing font chain for `calibri|0|0`');
          expect(fetch).not.toHaveBeenCalled();
          return;
        }
        const png = await exported([calibri]);
        expect(png).toBeInstanceOf(Uint8Array);
        expect(png.byteLength).toBeGreaterThan(0);
        const base = fonts.BASE_LATIN_FONTS.find(({ weight, style }) => weight === 400 && style === 'normal')!;
        expect(fetch.mock.calls.some(([url]) => String(url).endsWith(base.file))).toBe(true);
      } finally {
        fetch.mockRestore();
        provider.mockRestore();
        setFontCdnConfiguration(previous);
      }
    });

    test('T14: DOCX PNG reuses the browser worker layout until its CDN configuration changes', async () => {
      stubBrowser();
      const session = await createYrsSession();
      const previous = getFontCdnConfiguration();
      const restore = withBrowserFontGlobals(false);
      const provider = spyOn(fonts, 'createFontProvider').mockImplementation(fontRuntimeWithoutPacks().createFontProvider);
      const fetch = spyOnFontFetch(fontAssetResponse);
      try {
        session.openDocx(runsDocument({ family: 'Calibri', text: 'Plain text' }), true);
        fonts.configureFontCdn({ baseUrl: 'https://docx-png-first.example/' });
        await session.renderPng({ page: 0 });
        const providers = provider.mock.calls.length;
        await session.renderPng({ page: 0 });
        expect(provider).toHaveBeenCalledTimes(providers);
        fonts.configureFontCdn({ baseUrl: 'https://docx-png-second.example/' });
        await session.renderPng({ page: 0 });
        expect(fetch.mock.calls.some(([url]) => String(url).startsWith('https://docx-png-second.example/'))).toBe(true);
      } finally {
        session.destroy();
        fetch.mockRestore();
        provider.mockRestore();
        restore();
        setFontCdnConfiguration(previous);
      }
    });

    test('an installed pack that cannot serve a face rejects the export naming the face and the family', async () => {
      stubBrowser();
      const face = fonts.resolveMetricCompatFace('Arial', false, false)!;
      withheld.add(face.file);
      const error = await rejection(exported([{ family: 'Arial', text: 'Plain text' }]));
      expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      expect(error.message).toContain(`${face.family} for Arial could not load`);
      expect(error.message).toContain(face.file);
    });

    test('a script face an installed pack cannot serve rejects the export naming it', async () => {
      stubBrowser();
      const face = fonts.resolveScriptFallbackFace('arabic', false, false, 'Arial')!;
      withheld.add(face.file);
      const error = await rejection(exported([{ family: 'Arial', text: 'مرحبا بالعالم', rtl: true }]));
      expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      expect(error.message).toContain(`${face.family} for `);
      expect(error.message).toContain(face.file);
    });

    test('a bold run in a family without a bundled bold face rejects when its regular face cannot load', async () => {
      stubBrowser();
      const face = fonts.resolveMetricCompatFace('DM Serif Display', true, false)!;
      expect(face).toMatchObject({ family: 'DM Serif Display', weight: 400, style: 'normal' });
      withheld.add(face.file);
      const error = await rejection(exported([{ family: 'DM Serif Display', text: 'Heading', bold: true }]));
      expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      expect(error.message).toContain(`${face.family} for DM Serif Display could not load`);
      expect(error.message).toContain(face.file);
    });

    test('every export sharing a layout whose face failed rejects, and a later export lays out again', async () => {
      stubBrowser();
      const face = fonts.resolveMetricCompatFace('Arial', false, false)!;
      withheld.add(face.file);
      const session = await createYrsSession();
      try {
        session.openDocx(runsDocument({ family: 'Arial', text: 'Plain text' }), true);
        const results = await Promise.allSettled([session.renderPng({ page: 0 }), session.renderPng({ page: 0 })]);
        expect(results.map((result) => result.status)).toEqual(['rejected', 'rejected']);
        for (const result of results) {
          const error = (result as PromiseRejectedResult).reason as { message: string };
          expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
          expect(error.message).toContain(`${face.family} for Arial could not load`);
          expect(error.message).toContain(face.file);
        }
        withheld.clear();
        expect(await session.renderPng({ page: 0 })).toBeInstanceOf(Uint8Array);
      } finally { session.destroy(); }
    });

    test('an export that rejected for a face that failed to fetch is followed by one of the same version that succeeds', async () => {
      stubBrowser();
      const face = fonts.resolveMetricCompatFace('Arial', false, false)!;
      let failing = true;
      gateLatinFetches(async (file) => failing && file === face.file ? new Response(null, { status: 503 }) : latinBytes(file));
      const session = await createYrsSession();
      try {
        session.openDocx(runsDocument({ family: 'Arial', text: 'Plain text' }), true);
        const version = session.version();
        const error = await rejection(session.renderPng({ page: 0 }));
        expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
        expect(error.message).toContain(face.file);
        failing = false;
        expect(session.version()).toBe(version);
        const bytes = await session.renderPng({ page: 0 });
        expect(bytes).toBeInstanceOf(Uint8Array);
        expect(bytes.byteLength).toBeGreaterThan(0);
      } finally { session.destroy(); }
    });

    test('an explicit fontProvider lays out without a face it cannot load, as before', async () => {
      stubBrowser();
      const fontProvider = await resolveCompleteFontProvider();
      const face = arabicFace();
      withheld.add(face.file);
      expect(await exported(latinAndArabic, { fontProvider })).toBeInstanceOf(Uint8Array);
      expect([...asked]).toContain(face.file);
    });

    test('an explicit fontProvider exports Latin text when Arial\'s regular face cannot load, as before', async () => {
      stubBrowser();
      const fontProvider = await resolveCompleteFontProvider();
      const face = fonts.resolveMetricCompatFace('Arial', false, false)!;
      withheld.add(face.file);
      expect(await exported([{ family: 'Arial', text: 'Plain text' }], { fontProvider })).toBeInstanceOf(Uint8Array);
      expect([...asked]).toContain(face.file);
    });

    test('a configured source still loading when the configuration resets lays out without a face it cannot load', async () => {
      stubBrowser();
      const face = arabicFace();
      withheld.add(face.file);
      let entered!: () => void;
      const loadStarted = new Promise<void>((resolve) => { entered = resolve; });
      let release!: () => void;
      const gate = new Promise<void>((resolve) => { release = resolve; });
      configureDefaultFonts({ load: async () => { entered(); await gate; return fonts; } });
      try {
        const exporting = exported(latinAndArabic);
        await loadStarted;
        configureDefaultFonts({});
        release();
        expect(await exporting).toBeInstanceOf(Uint8Array);
        expect([...asked]).toContain(face.file);
      } finally { release(); configureDefaultFonts({}); }
    });

    test('an export aborted while the shared layout loads its faces rejects at once, and another export of the document completes', async () => {
      stubBrowser();
      let entered!: () => void;
      const loadStarted = new Promise<void>((resolve) => { entered = resolve; });
      let release!: () => void;
      const gate = new Promise<void>((resolve) => { release = resolve; });
      gateLatinFetches(async (file) => {
        entered();
        await gate;
        return latinBytes(file);
      });
      const session = await createYrsSession();
      try {
        session.openDocx(scriptDocument(), true);
        const controller = new AbortController();
        const aborted = session.renderPng({ page: 0 }, { signal: controller.signal });
        const other = session.renderPng({ page: 0 });
        await loadStarted;
        controller.abort(new Error('stopped'));
        await expect(aborted).rejects.toThrow('stopped');
        release();
        expect(await other).toBeInstanceOf(Uint8Array);
      } finally { release(); session.destroy(); }
    });
  });
}

import { afterEach, beforeAll, describe, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as fonts from '@betteroffice/fonts';
import { configureDefaultFonts, resolveCompleteFontProvider } from '../layout/measure';
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
      withheld.add(fonts.resolveMetricCompatFace('Arial', false, false)!.file);
      const session = await createYrsSession();
      try {
        session.openDocx(runsDocument({ family: 'Arial', text: 'Plain text' }), true);
        const results = await Promise.allSettled([session.renderPng({ page: 0 }), session.renderPng({ page: 0 })]);
        expect(results.map((result) => result.status === 'rejected' && (result.reason as { code?: string }).code))
          .toEqual(['font-unavailable', 'font-unavailable']);
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

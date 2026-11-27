import * as fonts from '@betteroffice/fonts';
import { fontAssetResponse, withBrowserFontGlobals } from '../../../fonts/src/fontAssets.test-support';
import { spyOnFontFetch } from '../../../fonts/src/fontFetch.test-support';
import { createFontRuntime } from '../../../fonts/src/fontRuntime';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import type { FontCdnDiagnostic } from '../../../../shared/fontCdnDiagnostics';
import { expect, spyOn, test } from 'bun:test';
import { BUNDLED_FONTS } from '@betteroffice/fonts';
import cases from '../../../../crates/betteroffice-fonts/data/script-fallback-chains.json';
import { textScripts } from '../../../../shared/office-session/textScripts';
import { defaultSheetFaces } from './enginePngFonts';

const arabicFaces = BUNDLED_FONTS.filter((face) => face.script === 'arabic');

test('Arabic in bold italic Arial uses Noto Naskh Arabic in PNG exports', async () => {
  const font = { family: 'Arial', bold: true, italic: true };
  const { fallbacks } = await defaultSheetFaces([{ font, scripts: textScripts('مرحبا') }]);
  const arabic = fallbacks(font).filter((face) => arabicFaces.some((known) => known.family === face.family));
  expect(arabic).toMatchObject([{ family: 'Noto Naskh Arabic', bold: false, italic: false }]);
});

for (const { family, bold, arabic } of cases) {
  test(`a ${family}${bold ? ' bold' : ''} cell falls back for Arabic to the face the native chain gives`, async () => {
    const font = { family, bold, italic: false };
    const { faces, fallbacks } = await defaultSheetFaces([{ font, scripts: ['arabic'] }]);
    const chain = [...faces.filter((face) => face.family === family), ...fallbacks(font)];
    const files = chain.flatMap((face) =>
      arabicFaces.filter((known) => known.byteLength === face.bytes.byteLength).map((known) => known.file));
    expect(files).toEqual([arabic]);
  });
}

const request = { font: { family: 'Calibri', bold: false, italic: false }, scripts: [] };

async function withoutPacks(run: () => Promise<void>, installedPack?: 'latin') {
  const runtime = createFontRuntime(async (pack) => ({
    cdn: getFontCdnConfiguration(),
    urls: pack === 'base' || pack === installedPack ? Object.fromEntries(Object.entries((await fonts.fontAssetUrls(pack))!)
      .map(([file, href]) => [file, () => new URL(href)])) : undefined,
  }));
  const installed = spyOn(fonts, 'bundledFaceInstalled').mockImplementation(runtime.bundledFaceInstalled);
  const load = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(runtime.loadBundledFontBytes);
  const previous = getFontCdnConfiguration();
  setFontCdnConfiguration(undefined);
  try { await run(); }
  finally {
    installed.mockRestore();
    load.mockRestore();
    setFontCdnConfiguration(previous);
  }
}

test('XLSX PNG faces reject on an implicit CDN failure instead of painting a substitute', async () => withoutPacks(async () => {
  const restore = withBrowserFontGlobals();
  const fetchAsset = globalThis.fetch;
  globalThis.fetch = ((input: RequestInfo | URL) => String(input).includes('cdn.jsdelivr.net')
    ? Promise.resolve(new Response(null, { status: 404 })) : fontAssetResponse(input)) as typeof fetch;
  try {
    await expect(defaultSheetFaces([request])).rejects.toMatchObject({
      code: 'font-unavailable', message: expect.stringContaining('HTTP 404'),
    });
  } finally { globalThis.fetch = fetchAsset; restore(); }
}));

test.each([
  { family: 'MS Gothic', text: 'q漢한', pack: 'cjk' },
  { family: 'Noto Sans Thai', text: 'qกກ', pack: 'intl' },
])('XLSX PNG faces for $family match the absent pack when the default CDN is blocked', ({ family, text, pack }) => withoutPacks(async () => {
  const restore = withBrowserFontGlobals();
  const fetch = spyOnFontFetch(async (input) => {
    if (String(input).includes('cdn.jsdelivr.net')) throw new TypeError('Failed to fetch');
    return fontAssetResponse(input);
  });
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const diagnostics: FontCdnDiagnostic[] = [];
  const dispatch = Object.getOwnPropertyDescriptor(globalThis, 'dispatchEvent');
  Object.defineProperty(globalThis, 'dispatchEvent', {
    configurable: true,
    value(event: CustomEvent<FontCdnDiagnostic>) {
      if (event.type === 'betteroffice:font-pack-diagnostic') diagnostics.push(event.detail);
      return true;
    },
  });
  const font = { family, bold: false, italic: false };
  const signature = ({ faces, fallbacks }: Awaited<ReturnType<typeof defaultSheetFaces>>) => {
    const summarize = ({ bytes, ...face }: (typeof faces)[number]) => ({ ...face, length: bytes.byteLength, hash: Bun.hash(bytes) });
    return { faces: faces.map(summarize), fallbacks: fallbacks(font).map(summarize) };
  };
  const select = () => defaultSheetFaces([{ font, scripts: textScripts(text) }]).then(signature);
  try {
    fonts.configureFontCdn({ baseUrl: false });
    const absent = await select();
    expect(absent.faces).toEqual([expect.objectContaining({
      family, length: fonts.resolveMetricCompatFace('Arial', false, false)!.byteLength,
    })]);
    expect(absent.fallbacks).toEqual([]);
    setFontCdnConfiguration(undefined);
    const diagnostic = expect.objectContaining({ pack, reason: 'cdn-unavailable' });
    await expect(select()).resolves.toEqual(absent);
    expect(diagnostics).toEqual([diagnostic]);
    await expect(select()).resolves.toEqual(absent);
    expect(fetch.mock.calls.filter(([url]) => String(url).includes('cdn.jsdelivr.net'))).toHaveLength(1);
    expect(warning).toHaveBeenCalledTimes(1);
    expect(diagnostics).toEqual([diagnostic, diagnostic]);

    fonts.configureFontCdn({ baseUrl: 'https://cdn.jsdelivr.net/npm/' });
    await expect(select()).rejects.toMatchObject({ code: 'font-unavailable', message: expect.stringContaining('Failed to fetch') });
  } finally {
    if (dispatch) Object.defineProperty(globalThis, 'dispatchEvent', dispatch);
    else Reflect.deleteProperty(globalThis, 'dispatchEvent');
    warning.mockRestore();
    fetch.mockRestore();
    restore();
  }
}, 'latin'));

test('XLSX PNG faces use Base Latin in headless mode without the Latin pack', async () => withoutPacks(async () => {
  const base = fonts.BASE_LATIN_FONTS.find(({ weight, style }) => weight === 400 && style === 'normal')!;
  await expect(defaultSheetFaces([request]).then(({ faces }) => faces.map(({ bytes }) => bytes.byteLength)))
    .resolves.toEqual([base.byteLength]);
}));

test('XLSX PNG faces use the CDN configured for each export', async () => withoutPacks(async () => {
  const fetchAsset = globalThis.fetch;
  const requests: string[] = [];
  globalThis.fetch = ((input: RequestInfo | URL) => {
    requests.push(String(input));
    return fontAssetResponse(input);
  }) as typeof fetch;
  try {
    fonts.configureFontCdn({ baseUrl: 'https://xlsx-png-first.example/' });
    await defaultSheetFaces([request]);
    fonts.configureFontCdn({ baseUrl: 'https://xlsx-png-second.example/' });
    await defaultSheetFaces([request]);
    expect(requests.some((url) => url.startsWith('https://xlsx-png-second.example/'))).toBe(true);
  } finally { globalThis.fetch = fetchAsset; }
}));

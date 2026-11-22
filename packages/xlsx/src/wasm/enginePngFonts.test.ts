import * as fonts from '@betteroffice/fonts';
import { fontAssetResponse, withBrowserFontGlobals } from '../../../fonts/src/test-assets';
import { createFontRuntime } from '../../../fonts/src/fontRuntime';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import { expect, spyOn, test } from 'bun:test';
import { BUNDLED_FONTS } from '@betteroffice/fonts';
import cases from '../../../../crates/betteroffice-fonts/data/script-fallback-chains.json';
import { defaultSheetFaces } from './enginePngFonts';

const arabicFaces = BUNDLED_FONTS.filter((face) => face.script === 'arabic');

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

async function withoutPacks(run: () => Promise<void>) {
  const runtime = createFontRuntime(async (pack) => ({
    cdn: getFontCdnConfiguration(),
    urls: pack === 'base' ? Object.fromEntries(Object.entries((await fonts.fontAssetUrls('base'))!)
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
    Reflect.deleteProperty(globalThis, Symbol.for('betteroffice.fonts.cdn-warned'));
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

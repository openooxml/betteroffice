import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import * as fonts from '@betteroffice/fonts';
import type { XlsxPngOptions, XlsxPngTarget } from '@betteroffice/xlsx';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../shared/office-session/fontCdn';
import { fontAssetResponse, fontRuntimeWithoutPacks, withBrowserFontGlobals } from '../../fonts/src/fontAssets.test-support';
import { spyOnFontFetch } from '../../fonts/src/fontFetch.test-support';
import { createXlsxCommandController } from './commands/createXlsxCommandStore';
import { createInThreadAsyncEditorApi } from './createInThreadAsyncEditorApi';
import { exportFaces, xlsxDefaultFontManager } from './defaultFonts';
import { editorExportFaces } from './useDefaultFonts';
import type { XlsxEditorApi } from './XlsxEditor.types';

afterEach(() => { mock.restore(); });

function bundled(load: () => Promise<ArrayBuffer>, byteLength = 3) {
  const fonts = {
    createFontProvider: () => ({
      resolve: (family: string) => family === 'Calibri' ? load : undefined,
      resolveFamily: (family: string) => family === 'Inter' ? load : undefined,
    }),
    resolveMetricCompatFace: (family: string, bold: boolean) => family === 'Calibri' ? { weight: bold ? 700 : 400, style: 'normal', byteLength } : undefined,
    resolveBundledFamilyFace: (family: string, _bold: boolean, italic: boolean) => family === 'Inter' ? { weight: 400, style: italic ? 'italic' : 'normal', byteLength } : undefined,
  };
  return spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue(fonts as never);
}

test('an export gets the faces the editor paints with, under the workbook family names', async () => {
  bundled(async () => new Uint8Array([1, 2, 3]).buffer);
  const faces = await exportFaces([
    { family: 'Calibri', bold: true, italic: false },
    { family: 'Inter', bold: false, italic: true },
    { family: 'Unbundled Sans', bold: false, italic: false },
  ], null);
  expect(faces).toEqual([
    { family: 'Calibri', bold: true, italic: false, bytes: new Uint8Array([1, 2, 3]) },
    { family: 'Inter', bold: false, italic: true, bytes: new Uint8Array([1, 2, 3]) },
  ]);
});

test('a face the editor would paint with but cannot fetch rejects instead of exporting a substitute', async () => {
  bundled(async () => { throw new Error('offline'); });
  await expect(exportFaces([{ family: 'Calibri', bold: false, italic: false }], null))
    .rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable', message: expect.stringContaining('Calibri') });
});

test('a face whose bytes do not match the manifest is left out, as the editor leaves it out', async () => {
  bundled(async () => new Uint8Array([1, 2]).buffer);
  expect(await exportFaces([{ family: 'Calibri', bold: false, italic: false }], null)).toEqual([]);
});

test('editors with default fonts off forward no faces and load no fonts', async () => {
  const load = bundled(async () => new Uint8Array([1, 2, 3]).buffer);
  expect(await editorExportFaces([{ family: 'Calibri', bold: false, italic: false }], null, false)).toBeUndefined();
  expect(await editorExportFaces([], null, undefined)).toBeUndefined();
  expect(load).not.toHaveBeenCalled();
});

test('T12: editor export rejects CDN failures and unavailable pack faces with CDN off', async () => {
  const runtime = fontRuntimeWithoutPacks();
  spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({ ...fonts, default: fonts, ...runtime });
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals();
  const fetch = spyOnFontFetch((input) =>
    String(input).includes('cdn.jsdelivr.net') ? Promise.resolve(new Response(null, { status: 404 })) : fontAssetResponse(input));
  const requests = ['Calibri', 'Noto Sans Thai', 'MS Gothic'].map((family) => ({ family, bold: false, italic: false }));
  try {
    setFontCdnConfiguration(undefined);
    for (const request of requests) {
      await expect(exportFaces([request], null)).rejects.toMatchObject({ code: 'font-unavailable' });
    }
    expect(fetch.mock.calls.filter(([url]) => String(url).includes('cdn.jsdelivr.net'))).toHaveLength(requests.length);
    fonts.configureFontCdn({ baseUrl: false });
    for (const request of requests) {
      await expect(exportFaces([request], null)).rejects.toMatchObject({
        code: 'font-unavailable', message: expect.stringContaining(request.family),
      });
    }
    expect(fetch.mock.calls.filter(([url]) => String(url).includes('cdn.jsdelivr.net'))).toHaveLength(requests.length);
  } finally {
    restore();
    setFontCdnConfiguration(previous);
  }
});

test('an editor export with no resolved faces leaves core font selection enabled', async () => {
  bundled(async () => new ArrayBuffer(3));
  const png = new Uint8Array([1]);
  const renderPng = mock(async (_target: XlsxPngTarget, _options?: XlsxPngOptions) => png);
  const sync = {
    commands: createXlsxCommandController().store,
    flushPendingInput: async () => {},
    handle: { renderPng, renderPngFonts: () => [{ family: 'Unbundled Sans', bold: false, italic: false }] },
  } as unknown as XlsxEditorApi;
  const api = createInThreadAsyncEditorApi(sync, false, () => {});
  await expect(api.renderPng({ sheet: 0, range: 'A1' })).resolves.toBe(png);
  expect(renderPng.mock.lastCall?.[1]?.fonts).toBeUndefined();
});

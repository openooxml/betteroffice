import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import * as fonts from '@betteroffice/fonts';
import type { XlsxPngOptions, XlsxPngTarget } from '@betteroffice/xlsx';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../shared/office-session/fontCdn';
import { fontAssetResponse, fontRuntimeWithoutPacks, withBrowserFontGlobals } from '../../fonts/src/fontAssets.test-support';
import { spyOnFontFetch } from '../../fonts/src/fontFetch.test-support';
import { createFontRuntime } from '../../fonts/src/fontRuntime';
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

test('editor export omits optional faces when the default CDN fails', async () => {
  const runtime = fontRuntimeWithoutPacks();
  spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({ ...fonts, default: fonts, ...runtime });
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals();
  const fetch = spyOnFontFetch((input) =>
    String(input).includes('cdn.jsdelivr.net') ? Promise.resolve(new Response(null, { status: 404 })) : fontAssetResponse(input));
  const requests = ['Calibri', 'Noto Sans Thai', 'MS Gothic'].map((family) => ({ family, bold: false, italic: false }));
  const [calibri, ...optional] = requests;
  try {
    setFontCdnConfiguration(undefined);
    await expect(exportFaces([calibri!], null)).rejects.toMatchObject({ code: 'font-unavailable' });
    await expect(exportFaces(optional, null)).resolves.toEqual([]);
    await expect(editorExportFaces(optional, null, undefined)).resolves.toBeUndefined();
    expect(fetch.mock.calls.filter(([url]) => String(url).includes('cdn.jsdelivr.net'))).toHaveLength(requests.length);
    fonts.configureFontCdn({ baseUrl: false });
    for (const family of ['Calibri', 'Noto Sans Arabic', 'Noto Sans Hebrew']) {
      await expect(exportFaces([{ family, bold: false, italic: false }], null)).rejects.toMatchObject({
        code: 'font-unavailable', message: expect.stringContaining(family),
      });
    }
    expect(fetch.mock.calls.filter(([url]) => String(url).includes('cdn.jsdelivr.net'))).toHaveLength(requests.length);
    fonts.configureFontCdn({ baseUrl: 'https://cdn.jsdelivr.net/npm/' });
    for (const request of optional) {
      await expect(exportFaces([request], null)).rejects.toMatchObject({
        code: 'font-unavailable', message: expect.stringContaining(request.family),
      });
    }
  } finally {
    restore();
    setFontCdnConfiguration(previous);
  }
});

test('editor PNG export omits absent optional faces with the CDN disabled', async () => {
  const runtime = createFontRuntime(async (pack) => {
    const urls = pack === 'latin' ? await fonts.fontAssetUrls(pack) : undefined;
    return {
      cdn: getFontCdnConfiguration(),
      urls: urls && Object.fromEntries(Object.entries(urls).map(([file, href]) => [file, () => new URL(href)])),
    };
  });
  spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({ ...fonts, default: fonts, ...runtime });
  const fetch = spyOnFontFetch(fontAssetResponse);
  const previous = getFontCdnConfiguration();
  const restore = withBrowserFontGlobals();
  const target = { sheet: 0, range: 'A1' };
  const optional = ['MS Gothic', 'Noto Sans Thai'].map((family) => ({ family, bold: false, italic: false }));
  const calibri = { family: 'Calibri', bold: false, italic: false };
  const png = new Uint8Array([1]);
  const renderPng = mock(async (_target: XlsxPngTarget, _options?: XlsxPngOptions) => png);
  const renderPngFonts = mock(() => optional);
  const sync = {
    commands: createXlsxCommandController().store,
    flushPendingInput: async () => {},
    handle: { renderPng, renderPngFonts },
  } as unknown as XlsxEditorApi;
  const api = createInThreadAsyncEditorApi(sync, false, () => {});
  try {
    fonts.configureFontCdn({ baseUrl: false });
    await expect(api.renderPng(target)).resolves.toBe(png);
    expect(renderPng.mock.lastCall).toEqual([target, { fonts: undefined, signal: undefined }]);
    await expect(exportFaces(optional, null)).resolves.toEqual([]);
    expect(fetch).not.toHaveBeenCalled();

    renderPngFonts.mockReturnValue([...optional, calibri]);
    await expect(api.renderPng(target)).resolves.toBe(png);
    expect(renderPng.mock.lastCall?.[1]?.fonts?.map(({ bytes, ...face }) => ({ ...face, byteLength: bytes.byteLength })))
      .toEqual([{ ...calibri, byteLength: fonts.resolveMetricCompatFace('Calibri', false, false)!.byteLength }]);
    expect(fetch.mock.calls.filter(([url]) => String(url).startsWith('https:'))).toHaveLength(0);
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

test('in-thread async exports respect the editor’s current default-font setting', async () => {
  const load = bundled(async () => new ArrayBuffer(3));
  const renderPng = mock(async (_target: XlsxPngTarget, _options?: XlsxPngOptions) => new Uint8Array([1]));
  const sync = {
    commands: createXlsxCommandController().store,
    flushPendingInput: async () => {},
    handle: { renderPng, renderPngFonts: () => [{ family: 'Calibri', bold: false, italic: false }] },
  } as unknown as XlsxEditorApi;
  let defaultFonts = false;
  const api = createInThreadAsyncEditorApi(sync, false, () => {}, undefined, () => defaultFonts);
  const target = { sheet: 0, range: 'A1' };
  await api.renderPng(target);
  expect(renderPng.mock.lastCall?.[1]?.fonts).toBeUndefined();
  expect(load).not.toHaveBeenCalled();
  defaultFonts = true;
  await api.renderPng(target);
  expect(renderPng.mock.lastCall?.[1]?.fonts?.map(({ family }) => family)).toEqual(['Calibri']);
});

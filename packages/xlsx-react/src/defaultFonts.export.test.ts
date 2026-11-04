import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import { exportFaces, xlsxDefaultFontManager } from './defaultFonts';
import { editorExportFaces } from './useDefaultFonts';

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

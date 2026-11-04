import type { PngRunFont } from '../../../../shared/office-session/pngFonts';
import type { WorkbookHandle } from './loader';
import type { XlsxPngOptions, XlsxPngTarget } from './pngTarget';

/**
 * A session's lazily loaded rasterizer, for runtimes with neither a canvas nor
 * a rasterizer in the engine module. It draws a PNG display list.
 */
export const pngRasterizers = new WeakMap<WorkbookHandle, (displayListJson: string) => Promise<Uint8Array>>();

/** The families and styles a canvas export of a target paints, for callers that supply faces. */
export const pngFontReaders = new WeakMap<WorkbookHandle, (target: XlsxPngTarget, options: XlsxPngOptions) => PngRunFont[]>();

export function workbookPngFonts(handle: WorkbookHandle, target: XlsxPngTarget, options: XlsxPngOptions = {}): PngRunFont[] {
  const read = pngFontReaders.get(handle);
  if (!read) throw new Error('Workbook PNG fonts are unavailable');
  return read(target, options);
}

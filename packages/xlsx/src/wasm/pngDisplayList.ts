import type { PngRunFont } from '../../../../shared/office-session/pngFonts';
import type { WorkbookHandle } from './loader';
import type { XlsxPngOptions, XlsxPngTarget } from './pngTarget';

/** The families and styles a canvas export of a target paints, for callers that supply faces. */
export const pngFontReaders = new WeakMap<WorkbookHandle, (target: XlsxPngTarget, options: XlsxPngOptions) => PngRunFont[]>();

export function workbookPngFonts(handle: WorkbookHandle, target: XlsxPngTarget, options: XlsxPngOptions = {}): PngRunFont[] {
  const read = pngFontReaders.get(handle);
  if (!read) throw new Error('Workbook PNG fonts are unavailable');
  return read(target, options);
}

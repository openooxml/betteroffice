import { canvasPngAvailable, canvasUnavailable, displayListToPng, sheetFonts } from '../render/png';
import { pngDisplayList, type ResolvedPngTarget } from './pngResolve';
import type { XlsxPngOptions } from './pngTarget';
import type { PngFontFace } from '../../../../shared/office-session/pngFonts';

/** What an engine needs to draw one PNG. The workbook handle builds it; each engine reads what it uses. */
export interface XlsxPngEngine {
  /** Runs `read` with `sheet` active and puts the previous sheet back. */
  onSheet<T>(sheet: number, read: () => T): T;
  /** The engine document: `pngDisplayListJson` and `rangePngDisplayListJson`. */
  doc: {
    pngDisplayListJson(viewportJson: string): string;
    rangePngDisplayListJson(argsJson: string): string;
  };
}

/** Which engine this module draws with; the native package swaps in its own. */
export const engineKind: 'wasm' | 'native' = 'wasm';

let defaultFaces: typeof import('./enginePngFonts') | undefined;

/**
 * Draws a resolved target through the canvas. Without `fonts` it paints with the bundled faces the native engine
 * draws the sheet's fonts with, loaded on first use.
 */
export async function renderPng(engine: XlsxPngEngine, target: ResolvedPngTarget, options: XlsxPngOptions): Promise<Uint8Array> {
  if (!canvasPngAvailable()) throw canvasUnavailable();
  const { list, ...size } = pngDisplayList(engine, target, options);
  let fonts: ReadonlyArray<PngFontFace> | undefined = options.fonts;
  if (fonts === undefined) {
    defaultFaces ??= await import('./enginePngFonts');
    fonts = await defaultFaces.defaultSheetFaces(sheetFonts(list), options.signal);
  }
  return displayListToPng(list, { ...size, signal: options.signal, fonts });
}

export default renderPng;

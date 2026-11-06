import { canvasPngAvailable, canvasUnavailable, displayListToPng } from '../render/png';
import { pngDisplayList, type ResolvedPngTarget } from './pngResolve';
import type { XlsxPngOptions } from './pngTarget';

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

/** Draws a resolved target through the canvas. */
export async function renderPng(engine: XlsxPngEngine, target: ResolvedPngTarget, options: XlsxPngOptions): Promise<Uint8Array> {
  if (!canvasPngAvailable()) throw canvasUnavailable();
  const { list, ...size } = pngDisplayList(engine, target, options);
  return displayListToPng(list, { ...size, signal: options.signal, fonts: options.fonts });
}

export default renderPng;

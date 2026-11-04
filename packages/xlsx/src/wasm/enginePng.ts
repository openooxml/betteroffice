import { canvasPngAvailable, canvasUnavailable, displayListToPng } from '../render/png';
import { cappedViewport, pngDisplayList, rangeArgs, type ResolvedPngTarget } from './pngResolve';
import type { XlsxPngOptions } from './pngTarget';

/** What an engine needs to draw one PNG. The workbook handle builds it; each engine reads what it uses. */
export interface XlsxPngEngine {
  /** Runs `read` with `sheet` active and puts the previous sheet back. */
  onSheet<T>(sheet: number, read: () => T): T;
  /** The engine document: `pngDisplayListJson`, `rangePngDisplayListJson`, and in raster builds `renderPng`, `renderRangePng`. */
  doc: {
    pngDisplayListJson(viewportJson: string): string;
    rangePngDisplayListJson(argsJson: string): string;
    renderPng?(viewportJson: string): Uint8Array;
    renderRangePng?(argsJson: string): Uint8Array;
  };
  /** Whether this build's wasm still carries the rasterizer, for runtimes without a canvas. */
  rasterAvailable: boolean;
  /** A separately loaded rasterizer for the same runtimes, when the session provides one. */
  rasterize?: (displayListJson: string) => Promise<Uint8Array>;
}

/** Which engine this module draws with; the native package swaps in its own. */
export const engineKind: 'wasm' | 'native' = 'wasm';

/** Draws a resolved target through the canvas, or the build's rasterizer where no canvas exists. */
export function renderPng(engine: XlsxPngEngine, target: ResolvedPngTarget, options: XlsxPngOptions): Promise<Uint8Array> {
  const scale = options.scale ?? 1;
  if (canvasPngAvailable()) {
    let painted: ReturnType<typeof pngDisplayList>;
    try { painted = pngDisplayList(engine, target, options); } catch (error) { return Promise.reject(error); }
    const { list, ...size } = painted;
    return displayListToPng(list, { ...size, signal: options.signal, fonts: options.fonts });
  }
  if (!engine.rasterAvailable && !engine.rasterize) return Promise.reject(canvasUnavailable());
  if ('viewport' in target) {
    if (scale !== 1) return Promise.reject(canvasUnavailable('A scaled viewport PNG needs a canvas here.'));
    const json = JSON.stringify(cappedViewport(target.viewport, options));
    if (engine.rasterAvailable) return Promise.resolve(engine.onSheet(target.sheet, () => engine.doc.renderPng!(json).slice()));
    return engine.rasterize!(engine.onSheet(target.sheet, () => engine.doc.pngDisplayListJson(json)));
  }
  const args = rangeArgs(target.range, options);
  if (engine.rasterAvailable) return Promise.resolve(engine.onSheet(target.sheet, () => engine.doc.renderRangePng!(args).slice()));
  return engine.rasterize!(engine.onSheet(target.sheet, () => engine.doc.rangePngDisplayListJson(args)));
}

export default renderPng;

import type { DisplayList } from '../display-list/types';
import { canvasPngAvailable, canvasUnavailable, displayListToPng } from '../render/png';
import type { Viewport } from './loader';
import { cappedExtent, type ResolvedPngTarget } from './pngResolve';
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
  const canvas = canvasPngAvailable();
  if (!canvas && !engine.rasterAvailable && !engine.rasterize) {
    return Promise.reject(canvasUnavailable());
  }
  if ('viewport' in target) {
    const viewport: Viewport = {
      ...target.viewport,
      width: cappedExtent(target.viewport.width, options.maxWidth, scale),
      height: cappedExtent(target.viewport.height, options.maxHeight, scale),
    };
    if (!canvas) {
      if (scale !== 1) return Promise.reject(canvasUnavailable('A scaled viewport PNG needs a canvas here.'));
      const json = JSON.stringify(viewport);
      if (engine.rasterAvailable) return Promise.resolve(engine.onSheet(target.sheet, () => engine.doc.renderPng!(json).slice()));
      return engine.rasterize!(engine.onSheet(target.sheet, () => engine.doc.pngDisplayListJson(json)));
    }
    const list = JSON.parse(engine.onSheet(target.sheet, () => engine.doc.pngDisplayListJson(JSON.stringify(viewport)))) as DisplayList;
    return displayListToPng(list, {
      dpr: scale, signal: options.signal,
      width: Math.max(1, Math.ceil(Math.fround(viewport.width * scale))),
      height: Math.max(1, Math.ceil(Math.fround(viewport.height * scale))),
    });
  }
  const args = JSON.stringify({ range: target.range, scale, maxWidth: options.maxWidth, maxHeight: options.maxHeight });
  if (!canvas && engine.rasterAvailable) return Promise.resolve(engine.onSheet(target.sheet, () => engine.doc.renderRangePng!(args).slice()));
  if (!canvas) return engine.rasterize!(engine.onSheet(target.sheet, () => engine.doc.rangePngDisplayListJson(args)));
  const list = JSON.parse(engine.onSheet(target.sheet, () => engine.doc.rangePngDisplayListJson(args))) as DisplayList;
  return displayListToPng(list, {
    dpr: 1, signal: options.signal,
    width: Math.max(1, Math.ceil(list.width)), height: Math.max(1, Math.ceil(list.height)),
  });
}

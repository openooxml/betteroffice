import type { SlideDisplayList } from '../types';
import { type CanvasImageResolver, paintSlide, sizeCanvasForSlide } from './canvas';

export interface SlideToPngOptions {
  /** Output scale, e.g. 2 for hidpi. */
  scale?: number;
  resolveImage?: CanvasImageResolver;
}

/**
 * Rasterize one slide to a PNG blob through the same canvas painter the editor
 * draws with, so an export matches what is on screen. The server-side twin is
 * the `betteroffice-pptx-raster` crate.
 */
export async function slideToPng(
  list: SlideDisplayList,
  options: SlideToPngOptions = {}
): Promise<Blob> {
  const { scale = 1, resolveImage } = options;
  if (!Number.isFinite(scale) || scale <= 0) throw new Error('scale must be finite and positive');
  const canvas = createCanvas(list, scale);
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('2d canvas context unavailable');
  await paintSlide(ctx as CanvasRenderingContext2D, list, 1, scale, { resolveImage });
  return encode(canvas);
}

type ExportCanvas = HTMLCanvasElement | OffscreenCanvas;

function createCanvas(list: SlideDisplayList, scale: number): ExportCanvas {
  const width = Math.round(list.width * scale);
  const height = Math.round(list.height * scale);
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height);
  if (typeof document === 'undefined') throw new Error('no canvas backend in this environment');
  const canvas = document.createElement('canvas');
  // The style properties `sizeCanvasForSlide` writes are inert off-screen, but
  // going through it keeps one rounding rule for both paths.
  sizeCanvasForSlide(canvas, list, 1, scale);
  return canvas;
}

function encode(canvas: ExportCanvas): Promise<Blob> {
  if ('convertToBlob' in canvas) return canvas.convertToBlob({ type: 'image/png' });
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('canvas produced no png'));
    }, 'image/png');
  });
}

import { PngExportError } from '../../../../../shared/office-session/pngExportError';
import type { BundledFontProvider } from '../measure';
import type { DisplayPage } from './displayList';

/** The rasterizer's largest PNG side and area, in pixels. */
export const MAX_PAGE_PNG_SIDE = 16_384;
export const MAX_PAGE_PNG_PIXELS = 16_777_216;

export type DocxPngTarget = { page: number };

/** A face a PNG export may also lay out with, as `openDocx` hosts register them. */
export interface DocxPngFontFace {
  family: string;
  bold?: boolean;
  italic?: boolean;
  bytes: Uint8Array;
}

export interface DocxPngOptions {
  scale?: number;
  signal?: AbortSignal;
  fonts?: ReadonlyArray<DocxPngFontFace>;
  /** The provider the editor showing this document measures with, where it has one. */
  fontProvider?: BundledFontProvider;
}

/** Whether this runtime can paint and encode a PNG on a canvas: one with a 2d context, not only a document. */
export function canvasPngAvailable(): boolean {
  if (typeof OffscreenCanvas === 'function') return true;
  if (typeof document === 'undefined') return false;
  try { return document.createElement('canvas').getContext('2d') !== null; } catch { return false; }
}

export function canvasUnavailable(): PngExportError {
  return new PngExportError('canvas-unavailable', 'PNG export needs a canvas, and this environment has none. Use @betteroffice/docx-native.');
}

export function pictureUnavailable(): PngExportError {
  return new PngExportError('canvas-unavailable', 'This page has SVG or metafile pictures, which decode only with a document. Export from the page, or use @betteroffice/docx-native.');
}

export function fontUnavailable(detail: string): PngExportError {
  return new PngExportError('font-unavailable', detail);
}

/** Checks a page export's scale and size the way every engine does, and returns the PNG size. */
export function pagePngRequest(
  page: Pick<DisplayPage, 'width' | 'height'>, options: Pick<DocxPngOptions, 'scale'>,
): { width: number; height: number; scale: number } {
  const scale = pngScale(options.scale);
  const width = Math.ceil(page.width * scale);
  const height = Math.ceil(page.height * scale);
  if (width > MAX_PAGE_PNG_SIDE || height > MAX_PAGE_PNG_SIDE || width * height > MAX_PAGE_PNG_PIXELS) {
    throw new PngExportError('canvas-too-large', `A ${width}x${height}px page is over the ${MAX_PAGE_PNG_SIDE}px side or ${MAX_PAGE_PNG_PIXELS} pixel limit; lower the scale.`);
  }
  return { width, height, scale };
}

export function pngScale(value: unknown): number {
  const scale = value ?? 1;
  if (typeof scale !== 'number' || !Number.isFinite(scale) || scale <= 0) throw new RangeError('scale must be finite and positive');
  return scale;
}

/** Checks a target the way the native renderer does. */
export function pngPage(target: unknown): number {
  if (typeof target !== 'object' || target === null || Object.keys(target).length !== 1 || !('page' in target)) {
    throw new TypeError('target must be { page } with a page index');
  }
  const page = (target as DocxPngTarget).page;
  if (!Number.isInteger(page) || page < 0) throw new TypeError('page must be a non-negative integer');
  return page;
}

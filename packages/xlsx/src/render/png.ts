import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { DisplayList, TextCmd } from '../display-list/types';
import { MAX_PNG_PIXELS, MAX_PNG_SIDE } from '../wasm/pngTarget';
import { fontString, paintDisplayList } from './canvas2d';

export interface DisplayListToPngOptions {
  /** Device pixels per display-list pixel. */
  dpr: number;
  /** PNG size in pixels, as the engine sized it. */
  width: number;
  height: number;
  signal?: AbortSignal;
}

type ExportCanvas = HTMLCanvasElement | OffscreenCanvas;

/** Whether this runtime can paint and encode a PNG on a canvas. */
export function canvasPngAvailable(): boolean {
  return typeof OffscreenCanvas === 'function' || typeof document !== 'undefined';
}

/** Paints a sheet display list with the editor's painter on a white canvas and encodes it. */
export async function displayListToPng(list: DisplayList, options: DisplayListToPngOptions): Promise<Uint8Array> {
  const { width, height, dpr } = options;
  if (width > MAX_PNG_SIDE || height > MAX_PNG_SIDE || width * height > MAX_PNG_PIXELS) {
    throw new PngExportError('canvas-too-large', `A ${width}x${height}px PNG is over the ${MAX_PNG_SIDE}px side or ${MAX_PNG_PIXELS} pixel limit; lower the scale or set maxWidth and maxHeight.`);
  }
  options.signal?.throwIfAborted();
  await loadFonts(list);
  options.signal?.throwIfAborted();
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) throw new PngExportError('canvas-unavailable', 'PNG export needs a 2d canvas context.');
  paintDisplayList(ctx, list, dpr);
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalCompositeOperation = 'destination-over';
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);
  ctx.restore();
  options.signal?.throwIfAborted();
  return new Uint8Array(await (await encode(canvas)).arrayBuffer());
}

function createCanvas(width: number, height: number): ExportCanvas {
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height);
  if (typeof document === 'undefined') {
    throw new PngExportError('canvas-unavailable', 'PNG export needs a canvas, and this environment has none. Use @betteroffice/xlsx-native.');
  }
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function encode(canvas: ExportCanvas): Promise<Blob> {
  if ('convertToBlob' in canvas) return canvas.convertToBlob({ type: 'image/png' });
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new PngExportError('canvas-too-large', 'The canvas produced no PNG.'));
    }, 'image/png');
  });
}

function fontSet(): FontFaceSet | undefined {
  return typeof document !== 'undefined' ? document.fonts : (globalThis as { fonts?: FontFaceSet }).fonts;
}

/**
 * Waits for every face the sheet paints with, as the editor names them, so the
 * export never paints a fallback while the real face is still loading.
 */
async function loadFonts(list: DisplayList): Promise<void> {
  const set = fontSet();
  if (!set) return;
  const fonts = new Map<string, string>();
  for (const cmd of list.commands) {
    if (cmd.op !== 'text') continue;
    const font = fontString(cmd as TextCmd);
    fonts.set(font, (cmd as TextCmd).fontFamily ?? 'sans-serif');
  }
  for (const [font, family] of fonts) {
    await set.load(font).catch(() => {
      throw new PngExportError('font-unavailable', `The font "${family}" could not load, so the sheet was not exported with a substitute.`);
    });
  }
}

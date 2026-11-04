import { canvasPngAvailable, canvasUnavailable, slideListToPng, type PptxPngBackground } from '../render/png';
import { decodeBrowserPresentationImage } from '../render/browserImage';
import type { PptxFontFace, SlideDisplayList } from '../types';

/** What an engine needs to draw one slide. The presentation handle builds it; each engine reads what it uses. */
export interface PptxPngEngine {
  /** The slide laid out for export, leaving hit testing on the slide it had. */
  displayList(slide: number): SlideDisplayList;
  /** A picture's display bytes, after TIFF and metafile conversion, or undefined. */
  media(assetId: string): Uint8Array | undefined;
  /** The faces layout measured with, by the engine's id for each. */
  fonts(): Array<PptxFontFace & { id: number }>;
  /** The engine document and renderer, for engines that render PNG themselves. */
  doc: unknown;
  renderer: unknown;
}

export interface PptxEnginePngOptions {
  scale?: number;
  background?: PptxPngBackground;
  maxShadowPixels?: number;
  signal?: AbortSignal;
  fonts?: ReadonlyArray<PptxFontFace>;
}

/** Which engine this module draws with; the native package swaps in its own. */
export const engineKind: 'wasm' | 'native' = 'wasm';

function assetIds(primitives: SlideDisplayList['primitives'], ids: Set<string>): Set<string> {
  for (const primitive of primitives) {
    if (primitive.kind === 'image' && primitive.assetId) ids.add(primitive.assetId);
    if (primitive.kind === 'chart' || primitive.kind === 'table') assetIds(primitive.primitives, ids);
  }
  return ids;
}

/** Paints a slide through the canvas, with the faces layout measured with and any the caller adds. */
export async function renderPng(engine: PptxPngEngine, target: { slide: number }, options: PptxEnginePngOptions): Promise<Uint8Array> {
  if (!canvasPngAvailable()) throw canvasUnavailable();
  const list = engine.displayList(target.slide);
  const images = new Map<string, CanvasImageSource | null>();
  try {
    for (const assetId of assetIds(list.primitives, new Set())) {
      options.signal?.throwIfAborted();
      const bytes = engine.media(assetId);
      images.set(assetId, bytes?.byteLength ? await decodeBrowserPresentationImage(bytes, 'image could not decode').catch(() => null) : null);
    }
    return await slideListToPng(list, {
      ...options, fonts: [...engine.fonts(), ...(options.fonts ?? [])], resolveImage: (assetId) => images.get(assetId) ?? null,
    });
  } finally {
    for (const image of images.values()) (image as { close?: () => void } | null)?.close?.();
  }
}

export default renderPng;

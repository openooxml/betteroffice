import { canvasPngAvailable, canvasUnavailable, pictureUnavailable, slideListToPng, type PptxPngBackground } from '../render/png';
import { browserPresentationImageBlob, decodeBrowserPresentationImage, needsElementDecode } from '../render/browserImage';
import type { PptxFontFace, SlideDisplayList } from '../types';

/** What an engine needs to draw one slide. The presentation handle builds it; each engine reads what it uses. */
export interface PptxPngEngine {
  /** The slide laid out for export, leaving hit testing on the slide it had. */
  displayList(slide: number): SlideDisplayList;
  /** A picture's display bytes, after TIFF and metafile conversion, or undefined. */
  media(assetId: string): Uint8Array | undefined;
  /** The faces layout measured with, by the engine's id for each. */
  fonts(): Array<PptxFontFace & { id: number; fallback?: boolean }>;
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

let fontLayout: typeof import('./enginePngFonts') | undefined;

/** Frees the export renderer an export kept for `document`. */
export function releasePng(document: object): void {
  fontLayout?.releaseDefaultFontLayout(document);
}

function decodesOnlyWithDocument(bytes: Uint8Array): boolean {
  try { return needsElementDecode(browserPresentationImageBlob(bytes)); } catch { return false; }
}

function assetIds(primitives: SlideDisplayList['primitives'], ids: Set<string>): Set<string> {
  for (const primitive of primitives) {
    if (primitive.kind === 'image' && primitive.assetId) ids.add(primitive.assetId);
    if (primitive.kind === 'chart' || primitive.kind === 'table') assetIds(primitive.primitives, ids);
  }
  return ids;
}

/**
 * Paints a slide through the canvas. With a document and the slide's ID it lays that slide out again with the faces the presentation
 * registered, any new family in `options.fonts`, and the bundled faces for the rest, as the native export does.
 */
export async function renderPng(engine: PptxPngEngine, target: { slide: number; id?: string }, options: PptxEnginePngOptions): Promise<Uint8Array> {
  if (!canvasPngAvailable()) throw canvasUnavailable();
  let list: SlideDisplayList;
  let fonts: ReadonlyArray<PptxFontFace>;
  if (engine.doc && target.id !== undefined) {
    const measured = engine.fonts();
    fontLayout ??= await import('./enginePngFonts');
    const layout = await fontLayout.defaultFontLayout(
      engine.doc as Parameters<typeof fontLayout.defaultFontLayout>[0], target.id, measured, options.fonts, options.signal);
    list = layout.list;
    fonts = [...layout.fonts, ...(options.fonts ?? [])];
  } else {
    list = engine.displayList(target.slide);
    fonts = [...engine.fonts(), ...(options.fonts ?? [])];
  }
  const images = new Map<string, CanvasImageSource | null>();
  try {
    for (const assetId of assetIds(list.primitives, new Set())) {
      options.signal?.throwIfAborted();
      const bytes = engine.media(assetId);
      if (bytes?.byteLength && typeof Image === 'undefined' && decodesOnlyWithDocument(bytes)) {
        throw pictureUnavailable('svg');
      }
      images.set(assetId, bytes?.byteLength ? await decodeBrowserPresentationImage(bytes, 'image could not decode').catch(() => null) : null);
    }
    return await slideListToPng(list, {
      ...options, fonts, resolveImage: (assetId) => images.get(assetId) ?? null,
    });
  } finally {
    for (const image of images.values()) (image as { close?: () => void } | null)?.close?.();
  }
}

export default renderPng;

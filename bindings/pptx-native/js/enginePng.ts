import { PngExportError } from '../../../shared/office-session/pngExportError';
import { slidePngRequest } from '../../../packages/pptx/src/render/png';
import type { PptxEnginePngOptions, PptxPngEngine } from '../../../packages/pptx/src/wasm/enginePng';
import type { PptxFontFace } from '../../../packages/pptx/src/types';
import { PptxRenderer } from './engine';

export type { PptxEnginePngOptions, PptxPngEngine } from '../../../packages/pptx/src/wasm/enginePng';

export const engineKind: 'wasm' | 'native' = 'native';

interface NativeRenderer {
  registerFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
  registerFallbackFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
  registerDefaultFonts(document: unknown, slide: number): void;
  renderPngTarget(document: unknown, targetJson: string, optionsJson: string): Uint8Array;
  free(): void;
}

/** One export renderer per document, so `layoutSlide` keeps measuring with the faces the caller registered. */
const exporters = new WeakMap<object, { renderer: NativeRenderer; faces: number }>();

/** Layout measured every style of a registered family with that family's faces, so an export cannot add to it. */
const familyKey = (face: PptxFontFace) => face.family.trim().toLowerCase();

/** Frees the document's export renderer; the presentation handle calls it on dispose. */
export function releasePng(document: object): void {
  exporters.get(document)?.renderer.free();
  exporters.delete(document);
}

function exportRenderer(faces: ReadonlyArray<PptxFontFace & { fallback?: boolean }>): NativeRenderer {
  const renderer = new PptxRenderer() as unknown as NativeRenderer;
  try {
    for (const face of faces) {
      const register = face.fallback ? renderer.registerFallbackFont : renderer.registerFont;
      register.call(renderer, face.family, face.bold ?? false, face.italic ?? false, face.bytes);
    }
    return renderer;
  } catch (error) {
    renderer.free();
    throw error;
  }
}

/**
 * Renders in the native engine on the calling thread, with the faces the
 * presentation registered, any in `options.fonts`, and the deck's embedded
 * faces, then the font packs for the rest. Options and size are checked as on
 * the canvas engine, so both reject alike.
 */
export async function renderPng(engine: PptxPngEngine, target: { slide: number }, options: PptxEnginePngOptions): Promise<Uint8Array> {
  const request = slidePngRequest(engine.displayList(target.slide), options);
  options.signal?.throwIfAborted();
  const document = engine.doc as object;
  const faces = engine.fonts();
  const measured = new Set(faces.map(familyKey));
  const added = options.fonts?.filter((face) => !measured.has(familyKey(face)));
  const supplied = added?.length ? added : undefined;
  let held = supplied ? undefined : exporters.get(document);
  if (held && held.faces !== faces.length) {
    held.renderer.free();
    exporters.delete(document);
    held = undefined;
  }
  const renderer = held?.renderer ?? exportRenderer(supplied ? [...faces, ...supplied] : faces);
  if (!supplied && !held) exporters.set(document, { renderer, faces: faces.length });
  let bytes: Uint8Array;
  try {
    renderer.registerDefaultFonts(document, target.slide);
    bytes = renderer.renderPngTarget(document, JSON.stringify({ slide: target.slide }),
      JSON.stringify({ scale: request.scale, background: request.background, maxShadowPixels: options.maxShadowPixels }));
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (message.startsWith('png-font-unavailable: ')) throw new PngExportError('font-unavailable', message.slice(22));
    if (message.startsWith('png-too-large: ')) throw new PngExportError('canvas-too-large', message.slice(15));
    throw error;
  } finally {
    if (supplied) renderer.free();
  }
  options.signal?.throwIfAborted();
  return bytes;
}

export default renderPng;

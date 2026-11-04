import { PngExportError } from '../../../shared/office-session/pngExportError';
import { slidePngRequest } from '../../../packages/pptx/src/render/png';
import type { PptxEnginePngOptions, PptxPngEngine } from '../../../packages/pptx/src/wasm/enginePng';

export type { PptxEnginePngOptions, PptxPngEngine } from '../../../packages/pptx/src/wasm/enginePng';

export const engineKind: 'wasm' | 'native' = 'native';

interface NativeRenderer {
  renderPngTarget(document: unknown, targetJson: string, optionsJson: string): Uint8Array;
}

/**
 * Renders in the native engine on the calling thread, with the faces the
 * presentation registered. Options and size are checked as on the canvas
 * engine, so both reject alike; `fonts` is for the browser and is ignored.
 */
export async function renderPng(engine: PptxPngEngine, target: { slide: number }, options: PptxEnginePngOptions): Promise<Uint8Array> {
  const request = slidePngRequest(engine.displayList(target.slide), options);
  options.signal?.throwIfAborted();
  let bytes: Uint8Array;
  try {
    bytes = (engine.renderer as NativeRenderer).renderPngTarget(engine.doc, JSON.stringify({ slide: target.slide }),
      JSON.stringify({ scale: request.scale, background: request.background, maxShadowPixels: options.maxShadowPixels }));
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (message.startsWith('png-font-unavailable: ')) throw new PngExportError('font-unavailable', message.slice(22));
    if (message.startsWith('png-too-large: ')) throw new PngExportError('canvas-too-large', message.slice(15));
    throw error;
  }
  options.signal?.throwIfAborted();
  return bytes;
}

export default renderPng;

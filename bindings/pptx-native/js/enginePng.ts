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
  const bytes = (engine.renderer as NativeRenderer).renderPngTarget(engine.doc, JSON.stringify({ slide: target.slide }),
    JSON.stringify({ scale: request.scale, background: request.background, maxShadowPixels: options.maxShadowPixels }));
  options.signal?.throwIfAborted();
  return bytes;
}

export default renderPng;

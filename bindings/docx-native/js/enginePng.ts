import { fontUnavailable, pagePngRequest, type DocxPngOptions } from '../../../packages/docx/src/layout/render/png';
import { PngExportError } from '../../../shared/office-session/pngExportError';
import type { DocxPngEngine } from '../../../packages/docx/src/wasm/enginePng';

export type { DocxPngEngine } from '../../../packages/docx/src/wasm/enginePng';

export const engineKind: 'wasm' | 'native' = 'native';

interface NativeSession {
  renderPngTarget(displayListJson: string, targetJson: string, optionsJson: string): Uint8Array;
}

/**
 * Rasterizes the page's display list in the native engine on the calling
 * thread, with the faces and media of the session that laid it out. Scale and
 * size are checked as on the canvas engine, so both reject alike.
 */
export async function renderPng(engine: DocxPngEngine, options: Pick<DocxPngOptions, 'scale' | 'signal'>): Promise<Uint8Array> {
  const { scale } = pagePngRequest(engine.page, options);
  options.signal?.throwIfAborted();
  let bytes: Uint8Array;
  try {
    bytes = (engine.session as NativeSession).renderPngTarget(
      JSON.stringify({ ...engine.displayList, pages: [engine.page] }), '{"page":0}', JSON.stringify({ scale }));
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (message.startsWith('png-font-unavailable: ')) throw fontUnavailable(message.slice(22));
    if (message.startsWith('png-too-large: ')) throw new PngExportError('canvas-too-large', message.slice(15));
    throw error;
  }
  options.signal?.throwIfAborted();
  return bytes;
}

export default renderPng;

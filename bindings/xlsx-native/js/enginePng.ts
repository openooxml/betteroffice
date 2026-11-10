import type { XlsxPngEngine } from '../../../packages/xlsx/src/wasm/enginePng';
import type { ResolvedPngTarget } from '../../../packages/xlsx/src/wasm/pngResolve';
import type { XlsxFontFace, XlsxPngOptions } from '../../../packages/xlsx/src/wasm/pngTarget';
import { fontUnavailable } from '../../../shared/office-session/pngFonts';

export type { XlsxPngEngine } from '../../../packages/xlsx/src/wasm/enginePng';

export const engineKind: 'wasm' | 'native' = 'native';

const UNAVAILABLE = 'png-font-unavailable: ';

interface NativePngDocument {
  renderPngTarget(targetJson: string, optionsJson: string, fonts?: ReadonlyArray<XlsxFontFace>): Uint8Array;
}

/**
 * Renders in the native engine on the calling thread; the Promise settles once
 * the PNG is encoded. Native errors pass through as they are. The native call
 * returns a fresh Uint8Array, so it is not copied again.
 */
export function renderPng(engine: XlsxPngEngine, target: ResolvedPngTarget, options: XlsxPngOptions): Promise<Uint8Array> {
  try {
    options.signal?.throwIfAborted();
    const doc = engine.doc as unknown as NativePngDocument;
    const { signal: _signal, fonts, ...native } = options;
    return Promise.resolve(doc.renderPngTarget(JSON.stringify(target), JSON.stringify(native), fonts));
  } catch (error) {
    const message = error instanceof Error ? error.message : typeof error === 'string' ? error : '';
    return Promise.reject(message.startsWith(UNAVAILABLE) ? fontUnavailable(message.slice(UNAVAILABLE.length), 'sheet') : error);
  }
}

export default renderPng;

import type { XlsxPngEngine } from '../../../packages/xlsx/src/wasm/enginePng';
import type { ResolvedPngTarget, XlsxPngOptions } from '../../../packages/xlsx/src/wasm/pngTarget';

export type { XlsxPngEngine } from '../../../packages/xlsx/src/wasm/enginePng';

interface NativePngDocument {
  renderPngTarget(targetJson: string, optionsJson: string): Uint8Array;
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
    const { signal: _signal, ...native } = options;
    return Promise.resolve(doc.renderPngTarget(JSON.stringify(target), JSON.stringify(native)));
  } catch (error) {
    return Promise.reject(error);
  }
}

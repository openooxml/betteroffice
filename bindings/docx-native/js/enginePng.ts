import { fontUnavailable, pagePngRequest, type DocxPngOptions } from '../../../packages/docx/src/layout/render/png';
import { DEFAULT_STRICT_PICTURES, PngExportError, strictPicturesError, type PngUnavailablePicture } from '../../../shared/office-session/pngExportError';
import type { DocxPngEngine } from '../../../packages/docx/src/wasm/enginePng';

export type { DocxPngEngine } from '../../../packages/docx/src/wasm/enginePng';

export const engineKind: 'wasm' | 'native' = 'native';

interface NativeSession {
  renderPngTargetReport(displayListJson: string, targetJson: string, optionsJson: string): { png: Uint8Array; pictures: PngUnavailablePicture[] };
}

/**
 * Rasterizes the page's display list in the native engine on the calling
 * thread, with the faces and media of the session that laid it out. Scale and
 * size are checked as on the canvas engine, so both reject alike.
 */
export async function renderPng(engine: DocxPngEngine, options: Pick<DocxPngOptions, 'scale' | 'signal' | 'onDiagnostic' | 'strictPictures'>): Promise<Uint8Array> {
  const { scale } = pagePngRequest(engine.page, options);
  options.signal?.throwIfAborted();
  let report: ReturnType<NativeSession['renderPngTargetReport']>;
  try {
    report = (engine.session as NativeSession).renderPngTargetReport(
      JSON.stringify({ ...engine.displayList, fontChains: engine.fontChains, pages: [engine.page] }), '{"page":0}', JSON.stringify({ scale }));
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (message.startsWith('png-font-unavailable: ')) throw fontUnavailable(message.slice(22));
    if (message.startsWith('png-too-large: ')) throw new PngExportError('canvas-too-large', message.slice(15));
    throw error;
  }
  options.signal?.throwIfAborted();
  if (report.pictures.length && (options.strictPictures ?? DEFAULT_STRICT_PICTURES)) throw strictPicturesError(report.pictures);
  if (report.pictures.length) {
    options.onDiagnostic?.({ code: 'picture-unavailable', pictures: report.pictures.map((picture) => ({ ...picture })) });
  }
  return report.png;
}

export default renderPng;

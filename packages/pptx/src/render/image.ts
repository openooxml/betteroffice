import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';
import { decodeTiffImage } from '../wasm/loader';
import { browserPresentationImageBlob, decodeBrowserImageBlob } from './browserImage';
import type { PresentationImageDecodeOptions } from './browserImage';

export { imageDecodeScale, needsElementDecode, setImageDecodeScale } from './browserImage';
export type { PresentationImageDecodeOptions } from './browserImage';

export function presentationImageBlob(bytes: Uint8Array): Blob {
  if (isTiff(bytes)) {
    if (bytes.byteLength > MAX_TIFF_BYTES) throw new Error('TIFF image exceeds the browser transfer budget');
    return new Blob([decodeTiffImage(bytes).slice()], { type: 'image/png' });
  }
  return browserPresentationImageBlob(bytes);
}

export async function decodePresentationImage(
  bytes: Uint8Array,
  errorMessage: string,
  options: PresentationImageDecodeOptions = {}
): Promise<CanvasImageSource> {
  const maximum = options.maxDimension;
  if (maximum != null && (!Number.isSafeInteger(maximum) || maximum < 1))
    throw new Error('invalid image dimension limit');
  return decodeBrowserImageBlob(presentationImageBlob(bytes), errorMessage, options);
}

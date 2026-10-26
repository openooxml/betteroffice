import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';
import { decodeTiffImage } from '../wasm/loader';
import { browserPresentationImageBlob, decodeBrowserImageBlob } from './browserImage';
import { metafileImageBytes } from './metafileImage';
import type { PresentationImageDecodeOptions as BrowserImageDecodeOptions } from './browserImage';

export { imageDecodeScale, needsElementDecode, setImageDecodeScale } from './browserImage';
export interface PresentationImageDecodeOptions extends BrowserImageDecodeOptions {
  metafileSvg?: () => Uint8Array | undefined;
}

export function presentationImageBlob(bytes: Uint8Array, metafileSvg?: () => Uint8Array | undefined): Blob {
  if (isTiff(bytes)) {
    if (bytes.byteLength > MAX_TIFF_BYTES) throw new Error('TIFF image exceeds the browser transfer budget');
    return new Blob([decodeTiffImage(bytes).slice()], { type: 'image/png' });
  }
  return browserPresentationImageBlob(metafileImageBytes(bytes, metafileSvg));
}

export async function decodePresentationImage(
  bytes: Uint8Array,
  errorMessage: string,
  options: PresentationImageDecodeOptions = {}
): Promise<CanvasImageSource> {
  const maximum = options.maxDimension;
  if (maximum != null && (!Number.isSafeInteger(maximum) || maximum < 1))
    throw new Error('invalid image dimension limit');
  const blob = presentationImageBlob(bytes, options.metafileSvg);
  if (blob.size === 0) throw new Error(errorMessage);
  return decodeBrowserImageBlob(blob, errorMessage, options);
}

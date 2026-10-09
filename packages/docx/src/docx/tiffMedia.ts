import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';

export interface TiffMediaResult {
  bytes: Uint8Array;
  mimeType: string;
  dataUrl: string;
}

/**
 * TIFF media to PNG via the parse wasm; non-TIFF passes through untouched.
 * The parse wasm already transcodes what it can, so anything still TIFF here is
 * an encoding it rejected and has warned about. Keep the source rather than
 * failing the open — a decoder that handles it still renders.
 */
export function maybeTranscodeTiffMedia(
  bytes: Uint8Array,
  mimeType: string,
  dataUrl: string,
  transcode: (bytes: Uint8Array) => Uint8Array
): TiffMediaResult {
  if (!isTiff(bytes) || bytes.byteLength > MAX_TIFF_BYTES) return { bytes, mimeType, dataUrl };
  try {
    const png = transcode(bytes);
    return {
      bytes: png,
      mimeType: 'image/png',
      dataUrl: `data:image/png;base64,${encodeBase64(png)}`,
    };
  } catch {
    return { bytes, mimeType, dataUrl };
  }
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

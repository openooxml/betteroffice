/** Browser TIFF transfer budget shared by the format image paths. */
export const MAX_TIFF_BYTES = 32 * 1024 * 1024;

/** TIFF magic-byte sniff shared by the browser image paths. */
export function isTiff(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    ((bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0) ||
      (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && bytes[3] === 0x2a))
  );
}

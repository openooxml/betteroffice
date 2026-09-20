import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';
import { decodeTiffImage } from '../wasm/loader';

const MAX_BITMAP_PIXELS = 33_554_432;

/** Convert presentation image formats that browsers cannot decode. */
export function presentationImageBlob(bytes: Uint8Array): Blob {
  if (isTiff(bytes)) {
    if (bytes.byteLength > MAX_TIFF_BYTES) {
      throw new Error('TIFF image exceeds the browser transfer budget');
    }
    return new Blob([decodeTiffImage(bytes).slice()], { type: 'image/png' });
  }
  const bitmap = wmfBitmap(bytes) ?? emfBitmap(bytes);
  return bitmap
    ? new Blob([bitmap], { type: 'image/bmp' })
    : new Blob([bytes.slice()]);
}

/** An EMF whose only ink is one unscaled `EMR_STRETCHDIBITS` filling its bounds. */
function emfBitmap(bytes: Uint8Array): Uint8Array<ArrayBuffer> | undefined {
  const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (
    data.byteLength < 108 ||
    data.getUint32(0, true) !== 1 ||
    data.getUint32(40, true) !== 0x464d4520
  ) return;
  const header = data.getUint32(4, true);
  if (header < 88 || header % 4 || header > data.byteLength - 8) return;
  const width = data.getInt32(16, true) - data.getInt32(8, true) + 1;
  const height = data.getInt32(20, true) - data.getInt32(12, true) + 1;
  const size = data.getUint32(header + 4, true);
  if (
    data.getUint32(header, true) !== 81 ||
    size < 80 ||
    size > data.byteLength - header - 8
  ) return;
  const at = (offset: number) => data.getInt32(header + offset, true);
  const bmi = data.getUint32(header + 48, true);
  const bits = data.getUint32(header + 56, true);
  const bitsBytes = data.getUint32(header + 60, true);
  if (
    at(24) !== data.getInt32(8, true) || at(28) !== data.getInt32(12, true) ||
    at(72) !== width || at(76) !== height ||
    at(32) !== 0 || at(36) !== 0 || at(40) !== width || at(44) !== height ||
    data.getUint32(header + 64, true) !== 0 ||
    data.getUint32(header + 68, true) !== 0x00cc0020 ||
    bmi < 80 || bmi + data.getUint32(header + 52, true) !== bits ||
    bits + bitsBytes !== size
  ) return;
  const end = header + size;
  if (
    data.getUint32(end, true) !== 14 ||
    data.getUint32(end + 4, true) !== data.byteLength - end
  ) return;
  return dibBitmap(bytes.subarray(header + bmi, end), width, height);
}

function wmfBitmap(bytes: Uint8Array): Uint8Array<ArrayBuffer> | undefined {
  const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (data.byteLength < 18) return;
  const start = data.getUint32(0, true) === 0x9ac6cdd7 ? 22 : 0;
  if (
    data.byteLength < start + 18 ||
    ![1, 2].includes(data.getUint16(start, true)) ||
    data.getUint16(start + 2, true) !== 9 ||
    data.getUint16(start + 4, true) !== 0x0300 ||
    data.getUint32(start + 6, true) * 2 !== data.byteLength - start
  ) return;
  let originX = 0;
  let originY = 0;
  let windowWidth = 0;
  let windowHeight = 0;
  let anisotropic = false;
  let bitmap: Uint8Array<ArrayBuffer> | undefined;
  for (let offset = start + 18; offset + 6 <= data.byteLength;) {
    const size = data.getUint32(offset, true) * 2;
    const command = data.getUint16(offset + 4, true);
    if (size < 6 || size > data.byteLength - offset) return;
    if (command === 0) return size === 6 && offset + size === data.byteLength ? bitmap : undefined;
    if (bitmap) return;
    if (command === 0x0103) {
      if (size < 8 || data.getUint16(offset + 6, true) !== 8) return;
      anisotropic = true;
    } else if (command === 0x0107) {
      if (size < 8 || ![1, 2, 3, 4].includes(data.getUint16(offset + 6, true))) return;
    } else if (command === 0x020b || command === 0x020c) {
      if (size !== 10) return;
      const y = data.getInt16(offset + 6, true);
      const x = data.getInt16(offset + 8, true);
      if (command === 0x020b) {
        originX = x;
        originY = y;
      } else {
        if (!anisotropic) return;
        windowWidth = x;
        windowHeight = y;
      }
    } else if (command === 0x0b41) {
      if (
        size < 66 ||
        data.getUint32(offset + 6, true) !== 0x00cc0020 ||
        data.getInt16(offset + 14, true) !== 0 ||
        data.getInt16(offset + 16, true) !== 0 ||
        data.getInt16(offset + 18, true) !== windowHeight ||
        data.getInt16(offset + 20, true) !== windowWidth ||
        data.getInt16(offset + 22, true) !== originY ||
        data.getInt16(offset + 24, true) !== originX ||
        !windowWidth || !windowHeight
      ) return;
      bitmap = dibBitmap(
        bytes.subarray(offset + 26, offset + size),
        data.getInt16(offset + 12, true),
        data.getInt16(offset + 10, true)
      );
      if (!bitmap) return;
    } else return;
    offset += size;
  }
}

function dibBitmap(
  bytes: Uint8Array,
  sourceWidth: number,
  sourceHeight: number
): Uint8Array<ArrayBuffer> | undefined {
  if (bytes.length < 40) return;
  const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = data.getInt32(4, true);
  const height = data.getInt32(8, true);
  const bits = data.getUint16(14, true);
  if (
    data.getUint32(0, true) !== 40 ||
    width < 1 || !height || width * Math.abs(height) > MAX_BITMAP_PIXELS ||
    sourceWidth !== width || sourceHeight !== Math.abs(height) ||
    data.getUint16(12, true) !== 1 ||
    ![24, 32].includes(bits) ||
    data.getUint32(16, true) !== 0 ||
    data.getUint32(32, true) !== 0
  ) return;
  const pixelBytes = Math.ceil(width * bits / 32) * 4 * Math.abs(height);
  const declaredBytes = data.getUint32(20, true);
  if (40 + pixelBytes !== bytes.length || (declaredBytes !== 0 && declaredBytes !== pixelBytes)) return;
  const bitmap = new Uint8Array(14 + bytes.length);
  const header = new DataView(bitmap.buffer);
  header.setUint16(0, 0x4d42, true);
  header.setUint32(2, bitmap.length, true);
  header.setUint32(10, 54, true);
  bitmap.set(bytes, 14);
  return bitmap;
}

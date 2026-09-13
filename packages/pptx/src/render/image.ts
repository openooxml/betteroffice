const MAX_BITMAP_PIXELS = 33_554_432;

/** Decode bitmap-only WMF wrappers without changing stored media. */
export function presentationImageBlob(bytes: Uint8Array): Blob {
  const bitmap = wmfBitmap(bytes);
  return bitmap
    ? new Blob([bitmap], { type: 'image/bmp' })
    : new Blob([bytes.slice()]);
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

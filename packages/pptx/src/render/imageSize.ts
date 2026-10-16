export interface ImageSize {
  width: number;
  height: number;
}

export function rasterImageSize(bytes: Uint8Array): ImageSize | undefined {
  const data = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 10) return;
  let width = 0;
  let height = 0;
  if (bytes.length >= 24 && data.getUint32(0) === 0x89504e47 && data.getUint32(4) === 0x0d0a1a0a && data.getUint32(12) === 0x49484452) {
    width = data.getUint32(16);
    height = data.getUint32(20);
  } else if (data.getUint32(0) === 0x47494638 && [0x37, 0x39].includes(bytes[4]) && bytes[5] === 0x61) {
    width = data.getUint16(6, true);
    height = data.getUint16(8, true);
  } else if (bytes.length >= 26 && data.getUint16(0, true) === 0x4d42) {
    if (data.getUint32(14, true) === 12) {
      width = data.getUint16(18, true);
      height = data.getUint16(20, true);
    } else {
      width = data.getInt32(18, true);
      height = Math.abs(data.getInt32(22, true));
    }
  } else if (bytes.length >= 20 && data.getUint32(0) === 0x52494646 && data.getUint32(8) === 0x57454250) {
    const chunk = data.getUint32(12);
    if (chunk === 0x56503858 && bytes.length >= 30) {
      width = 1 + bytes[24] + (bytes[25] << 8) + (bytes[26] << 16);
      height = 1 + bytes[27] + (bytes[28] << 8) + (bytes[29] << 16);
    } else if (chunk === 0x5650384c && bytes.length >= 25 && bytes[20] === 0x2f) {
      const bits = data.getUint32(21, true);
      width = 1 + (bits & 0x3fff);
      height = 1 + ((bits >>> 14) & 0x3fff);
    } else if (chunk === 0x56503820 && bytes.length >= 30 && bytes[23] === 0x9d && bytes[24] === 0x01 && bytes[25] === 0x2a) {
      width = data.getUint16(26, true) & 0x3fff;
      height = data.getUint16(28, true) & 0x3fff;
    }
  } else if (data.getUint16(0) === 0xffd8) {
    let rotated = false;
    for (let offset = 2; offset + 4 <= bytes.length;) {
      if (bytes[offset++] !== 0xff) return;
      while (bytes[offset] === 0xff) offset += 1;
      const marker = bytes[offset++];
      if (marker === 0xd9 || marker === 0xda) break;
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) continue;
      if (offset + 2 > bytes.length) return;
      const length = data.getUint16(offset);
      if (length < 2 || offset + length > bytes.length) return;
      if (marker === 0xe1) rotated = rotated || exifRotated(bytes.subarray(offset + 2, offset + length));
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker) && length >= 7) {
        height = data.getUint16(offset + 3);
        width = data.getUint16(offset + 5);
      }
      offset += length;
    }
    if (rotated) [width, height] = [height, width];
  }
  return width > 0 && height > 0 ? { width, height } : undefined;
}

function exifRotated(bytes: Uint8Array): boolean {
  if (bytes.length < 14 || ![0x45, 0x78, 0x69, 0x66, 0, 0].every((byte, index) => bytes[index] === byte)) return false;
  const data = new DataView(bytes.buffer, bytes.byteOffset + 6, bytes.byteLength - 6);
  const little = data.getUint16(0) === 0x4949;
  if ((!little && data.getUint16(0) !== 0x4d4d) || data.getUint16(2, little) !== 42) return false;
  const offset = data.getUint32(4, little);
  if (offset > data.byteLength - 2) return false;
  const count = data.getUint16(offset, little);
  for (let index = 0; index < count; index += 1) {
    const at = offset + 2 + index * 12;
    if (at > data.byteLength - 12) return false;
    if (data.getUint16(at, little) === 0x0112 && data.getUint16(at + 2, little) === 3 && data.getUint32(at + 4, little) === 1)
      return [5, 6, 7, 8].includes(data.getUint16(at + 8, little));
  }
  return false;
}

// intrinsic size of a PNG, GIF, JPEG, WebP or SVG, so post images reserve their space before they load
export function imageSize(bytes: Uint8Array): [number, number] | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const ascii = (at: number, length: number) => String.fromCharCode(...bytes.subarray(at, at + length));
  if (bytes.length >= 24 && ascii(1, 3) === "PNG") return [view.getUint32(16), view.getUint32(20)];
  if (bytes.length >= 10 && ascii(0, 3) === "GIF") return [view.getUint16(6, true), view.getUint16(8, true)];
  if (bytes.length >= 30 && ascii(0, 4) === "RIFF" && ascii(8, 4) === "WEBP") {
    const chunk = ascii(12, 4);
    const u24 = (at: number) => bytes[at] | (bytes[at + 1] << 8) | (bytes[at + 2] << 16);
    if (chunk === "VP8X") return [1 + u24(24), 1 + u24(27)];
    if (chunk === "VP8L") {
      const bits = view.getUint32(21, true);
      return [1 + (bits & 0x3fff), 1 + ((bits >> 14) & 0x3fff)];
    }
    if (chunk === "VP8 ") return [view.getUint16(26, true) & 0x3fff, view.getUint16(28, true) & 0x3fff];
    return null;
  }
  if (bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let at = 2;
    while (at + 9 < bytes.length) {
      if (bytes[at] !== 0xff) return null;
      const marker = bytes[at + 1];
      const length = view.getUint16(at + 2);
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return [view.getUint16(at + 7), view.getUint16(at + 5)];
      }
      at += 2 + length;
    }
    return null;
  }
  const text = new TextDecoder().decode(bytes.subarray(0, 4096));
  const svg = text.match(/<svg\b[^>]*>/);
  if (!svg) return null;
  const attr = (name: string) => svg[0].match(new RegExp(`\\b${name}="([\\d.]+)(?:px)?"`))?.[1];
  const [width, height] = [attr("width"), attr("height")].map(Number);
  if (width > 0 && height > 0) return [width, height];
  const box = svg[0].match(/\bviewBox="[\d.-]+[\s,]+[\d.-]+[\s,]+([\d.]+)[\s,]+([\d.]+)"/);
  return box ? [Number(box[1]), Number(box[2])] : null;
}

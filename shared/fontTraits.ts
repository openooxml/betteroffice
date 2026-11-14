interface FontTraits {
  weight: number;
  bold: boolean;
  italic: boolean;
  oblique: boolean;
}

export function readFontTraits(bytes: Uint8Array): FontTraits | undefined {
  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (view.byteLength < 12) return;
    let start = 0;
    if (view.getUint32(0) === 0x74746366) {
      const version = view.getUint32(4);
      const count = view.getUint32(8);
      if ((version !== 0x00010000 && version !== 0x00020000) || count === 0 ||
        12 + count * 4 + (version === 0x00020000 ? 12 : 0) > view.byteLength) return;
      start = view.getUint32(12);
    }
    if (start + 12 > view.byteLength) return;
    const signature = view.getUint32(start);
    if (signature !== 0x00010000 && signature !== 0x74727565 && signature !== 0x4f54544f) return;
    const count = view.getUint16(start + 4);
    if (start + 12 + count * 16 > view.byteLength) return;
    let os2: { offset: number; length: number } | undefined;
    let head: { offset: number; length: number } | undefined;
    for (let index = 0; index < count; index++) {
      const entry = start + 12 + index * 16;
      const tag = view.getUint32(entry);
      const offset = view.getUint32(entry + 8);
      const length = view.getUint32(entry + 12);
      if (offset + length > view.byteLength) return;
      if (tag === 0x4f532f32) {
        if (os2) return;
        os2 = { offset, length };
      } else if (tag === 0x68656164) {
        if (head) return;
        head = { offset, length };
      }
    }
    if (os2) {
      if (os2.length < 64) return;
      const selection = view.getUint16(os2.offset + 62);
      return {
        weight: Math.max(1, Math.min(1000, view.getUint16(os2.offset + 4))),
        bold: (selection & 32) !== 0,
        italic: (selection & 1) !== 0,
        oblique: (selection & 512) !== 0,
      };
    }
    if (head) {
      if (head.length < 46) return;
      const style = view.getUint16(head.offset + 44);
      return { weight: style & 1 ? 700 : 400, bold: (style & 1) !== 0, italic: (style & 2) !== 0, oblique: false };
    }
  } catch {
    return;
  }
}

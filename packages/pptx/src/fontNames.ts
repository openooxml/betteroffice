export function declaredFontFamilies(bytes: Uint8Array): ReadonlySet<string> {
  const families = new Set<string>();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.byteLength < 12) return families;
  const count = view.getUint16(4);
  if (12 + count * 16 > view.byteLength) return families;
  for (let table = 12; table < 12 + count * 16; table += 16) {
    if (view.getUint32(table) !== 0x6e616d65) continue;
    const start = view.getUint32(table + 8);
    const length = view.getUint32(table + 12);
    if (length < 6 || start + length > view.byteLength) return families;
    const records = view.getUint16(start + 2);
    const storage = start + view.getUint16(start + 4);
    if (storage < start + 6 + records * 12 || storage > start + length) return families;
    for (let record = start + 6; record < start + 6 + records * 12; record += 12) {
      const name = view.getUint16(record + 6);
      if (name !== 1 && name !== 16) continue;
      const platform = view.getUint16(record);
      const size = view.getUint16(record + 8);
      const offset = storage + view.getUint16(record + 10);
      if (offset + size > start + length) continue;
      let value = '';
      if (platform === 0 || platform === 3) {
        if (size % 2) continue;
        for (let index = offset; index < offset + size; index += 2) value += String.fromCharCode(view.getUint16(index));
      } else if (platform === 1) {
        for (let index = offset; index < offset + size; index++) value += String.fromCharCode(view.getUint8(index));
      }
      const family = value.trim().toLowerCase();
      if (family) families.add(family);
    }
    break;
  }
  return families;
}

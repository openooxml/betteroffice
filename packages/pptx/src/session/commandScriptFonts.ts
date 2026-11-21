import type { BundledFontScript } from '@betteroffice/fonts';
import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';

function hasFamily(bytes: ArrayBuffer, family: string): boolean {
  const view = new DataView(bytes);
  if (view.byteLength < 12) return false;
  const count = view.getUint16(4);
  if (12 + count * 16 > view.byteLength) return false;
  for (let table = 12; table < 12 + count * 16; table += 16) {
    if (view.getUint32(table) !== 0x6e616d65) continue;
    const start = view.getUint32(table + 8);
    const length = view.getUint32(table + 12);
    if (length < 6 || start + length > view.byteLength) return false;
    const records = view.getUint16(start + 2);
    const storage = start + view.getUint16(start + 4);
    if (6 + records * 12 > length) return false;
    for (let record = start + 6; record < start + 6 + records * 12; record += 12) {
      const name = view.getUint16(record + 6);
      if (name !== 1 && name !== 16) continue;
      const platform = view.getUint16(record);
      const size = view.getUint16(record + 8);
      const offset = storage + view.getUint16(record + 10);
      if (offset < start || offset + size > start + length) continue;
      let value = '';
      if (platform === 0 || platform === 3) {
        if (size % 2) continue;
        for (let index = offset; index < offset + size; index += 2) value += String.fromCharCode(view.getUint16(index));
      } else if (platform === 1) {
        for (let index = offset; index < offset + size; index++) value += String.fromCharCode(view.getUint8(index));
      }
      if (value.trim().toLowerCase() === family.trim().toLowerCase()) return true;
    }
  }
  return false;
}

export async function scriptFontFace(
  script: string, bold: boolean, italic: boolean, bytes: ArrayBuffer,
): Promise<DefaultFontFace | undefined> {
  const { resolveScriptFallbackFace } = await import('@betteroffice/fonts');
  const face = resolveScriptFallbackFace(script as BundledFontScript, bold, italic)
    ?? resolveScriptFallbackFace(script as BundledFontScript, false, false);
  if (!face || face.script !== script || !hasFamily(bytes, face.family)) return undefined;
  return { family: face.family, bold: face.weight === 700, italic: face.style === 'italic', bytes, kind: 'coverage' };
}

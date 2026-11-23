import type { BundledFontScript } from '@betteroffice/fonts';
import { importFonts } from '../../../../shared/office-session/fontModules';
import type { DefaultFontFace } from '../../../../shared/pptxDefaultFonts';
import { declaredFontFamilies } from '../fontNames';

export async function scriptFontFace(
  script: string, bold: boolean, italic: boolean, bytes: ArrayBuffer,
): Promise<DefaultFontFace | undefined> {
  const { resolveScriptFallbackFace } = await importFonts();
  const face = resolveScriptFallbackFace(script as BundledFontScript, bold, italic)
    ?? resolveScriptFallbackFace(script as BundledFontScript, false, false);
  if (!face || face.script !== script || !declaredFontFamilies(new Uint8Array(bytes)).has(face.family.trim().toLowerCase())) return undefined;
  return { family: face.family, bold: face.weight === 700, italic: face.style === 'italic', bytes, kind: 'coverage' };
}

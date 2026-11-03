import { loadBundledFontBytes } from '@betteroffice/fonts';
import type { PptxFontFace } from '@betteroffice/pptx';
import { preparePresentationFontDecoder } from '@betteroffice/pptx/worker';
import type { PptxBundledFont } from './bundledFonts';

/** Prepares the session worker for `file` and loads the bundled faces with its decoder, keeping input order. */
export function loadBundledFonts(file: Uint8Array, fonts: readonly (PptxFontFace | PptxBundledFont)[], warn: () => void) {
  const prepared = preparePresentationFontDecoder(file);
  const { decoder } = prepared;
  const loaded = Promise.all(fonts.map(async (font): Promise<PptxFontFace | undefined> => {
    if (!('face' in font)) return font;
    try {
      const bytes = await loadBundledFontBytes(font.face, { baseUrl: font.baseUrl }, { decoder });
      return { family: font.face.family, bold: font.face.weight === 700,
        italic: font.face.style === 'italic', bytes: new Uint8Array(bytes) };
    } catch {
      warn();
      return undefined;
    }
  }));
  return {
    fonts: loaded.then((list) => list.filter((font): font is PptxFontFace => font !== undefined)),
    release: () => prepared.release(),
  };
}

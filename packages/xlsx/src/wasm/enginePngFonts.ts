import {
  loadBundledFontBytes, resolveBundledFamilyFace, resolveGlyphFallbackFace, resolveLastResortChoice,
  type BundledFontFace,
} from '@betteroffice/fonts';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { PngRunFont } from '../../../../shared/office-session/pngFonts';
import type { XlsxFontFace } from './pngTarget';
import { decodeFontFace } from './loader';

const loaded = new Map<string, Promise<Uint8Array>>();

function load(face: BundledFontFace): Promise<Uint8Array> {
  let pending = loaded.get(face.file);
  if (!pending) {
    const started = loadBundledFontBytes(face, undefined, { decoder: decodeFontFace }).then((buffer) => new Uint8Array(buffer));
    started.catch(() => { if (loaded.get(face.file) === started) loaded.delete(face.file); });
    loaded.set(face.file, started);
    pending = started;
  }
  return pending;
}

/**
 * The face the native engine draws a cell font's Latin text with: the first face its font manager resolves the
 * family to, in its order (the family's bundled face, its glyph fallback, then Word's last resort).
 */
function firstFace(font: PngRunFont): BundledFontFace {
  const primary = resolveBundledFamilyFace(font.family, font.bold, font.italic);
  if (primary) return primary;
  const glyph = resolveGlyphFallbackFace(font.family, font.bold, font.italic, 'word');
  return glyph ?? resolveLastResortChoice(font.family, font.bold, font.italic, 'word').face;
}

/**
 * Bundled faces for the fonts a sheet's text paints with, as the native engine draws them, so a canvas export
 * without `fonts` matches it. Each face keeps its own weight and style, so a family drawn with a regular face, such
 * as bold Calibri Light, gets synthetic bold as natively. A face that cannot load rejects with `font-unavailable`
 * naming it; an abort rejects at once without waiting for loads other exports may share.
 */
export async function defaultSheetFaces(fonts: readonly PngRunFont[], signal?: AbortSignal): Promise<XlsxFontFace[]> {
  signal?.throwIfAborted();
  const loading = Promise.all(fonts.map(async (font) => {
    const face = firstFace(font);
    try {
      return { family: font.family, bold: face.weight === 700, italic: face.style === 'italic', bytes: await load(face) };
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      throw new PngExportError('font-unavailable', `${face.family} for ${font.family} could not load: ${reason}`);
    }
  }));
  if (!signal) return loading;
  let stop: (() => void) | undefined;
  const aborted = new Promise<never>((_, reject) => {
    stop = () => reject(signal.reason);
    signal.addEventListener('abort', stop, { once: true });
  });
  loading.catch(() => {});
  try {
    return await Promise.race([loading, aborted]);
  } finally {
    signal.removeEventListener('abort', stop!);
  }
}

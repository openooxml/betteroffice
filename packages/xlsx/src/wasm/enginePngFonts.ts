import type { BundledFontFace } from '@betteroffice/fonts';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { PngFontFace, PngRunFont } from '../../../../shared/office-session/pngFonts';
import { decodeFontFace } from './loader';

type Fonts = typeof import('@betteroffice/fonts');

let fontsModule: Promise<Fonts> | undefined;
/** The font packages, loaded with the first export that needs them; a worker resolves them only this way. */
const loadFonts = (): Promise<Fonts> => (fontsModule ??= import('@betteroffice/fonts').catch((error: unknown) => {
  fontsModule = undefined;
  throw error;
}));

const loaded = new Map<string, Promise<Uint8Array>>();

function load(fonts: Fonts, face: BundledFontFace): Promise<Uint8Array> {
  let pending = loaded.get(face.file);
  if (!pending) {
    const started = fonts.loadBundledFontBytes(face, undefined, { decoder: decodeFontFace, strict: true })
      .then((buffer) => new Uint8Array(buffer));
    started.catch(() => { if (loaded.get(face.file) === started) loaded.delete(face.file); });
    loaded.set(face.file, started);
    pending = started;
  }
  return pending;
}

/** An installed face, or undefined for one whose package is not installed. */
async function installedFace(fonts: Fonts, face: BundledFontFace | undefined): Promise<BundledFontFace | undefined> {
  return face && await fonts.bundledFaceInstalled(face).catch(() => true) ? face : undefined;
}

/**
 * The face the native engine draws a cell font's Latin text with: the first installed face its font manager
 * resolves the family to, in its order (the family's bundled face, its glyph fallback, then Word's last resort).
 */
async function firstFace(fonts: Fonts, font: PngRunFont): Promise<BundledFontFace> {
  const primary = await installedFace(fonts, fonts.resolveBundledFamilyFace(font.family, font.bold, font.italic));
  if (primary) return primary;
  const glyph = await installedFace(fonts, fonts.resolveGlyphFallbackFace(font.family, font.bold, font.italic, 'word'));
  return glyph ?? fonts.resolveLastResortChoice(font.family, font.bold, font.italic, 'word').face;
}

/**
 * Bundled faces for the fonts a sheet's text paints with, as the native engine draws them, so a canvas export
 * without `fonts` matches it. Each face serves the style the text asks for and installs with its own weight and
 * style, so bold Calibri Light drawn with a regular face gets synthetic bold, and regular Arial Black its bold face,
 * as natively. A face that cannot load rejects with `font-unavailable` naming it; an abort rejects at once without
 * waiting for loads other exports may share.
 */
export async function defaultSheetFaces(fonts: readonly PngRunFont[], signal?: AbortSignal): Promise<PngFontFace[]> {
  signal?.throwIfAborted();
  const loading = loadFonts().then((manager) => Promise.all(fonts.map(async (font) => {
    const face = await firstFace(manager, font);
    try {
      return {
        family: font.family, bold: font.bold, italic: font.italic, bytes: await load(manager, face),
        face: { bold: face.weight === 700, italic: face.style === 'italic' },
      };
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      throw new PngExportError('font-unavailable', `${face.family} for ${font.family} could not load: ${reason}`);
    }
  })));
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

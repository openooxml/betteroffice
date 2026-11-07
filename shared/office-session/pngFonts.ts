import { PngExportError } from './pngExportError';

/** A font face a PNG export may paint with. */
export interface PngFontFace {
  family: string;
  bold?: boolean;
  italic?: boolean;
  bytes: Uint8Array;
  /** The face's own weight and style, where it serves a style it is not, such as a regular face for bold text. */
  face?: { bold: boolean; italic: boolean };
}

/** A family and style a page or sheet asks for. */
export interface PngRunFont {
  family: string;
  bold: boolean;
  italic: boolean;
  /** The engine's id for the face layout measured with, where the engine reports one. */
  id?: number;
}

export interface InstalledPngFaces {
  /** The private family the export paints `font` with, when a supplied face covers it. */
  family(font: PngRunFont): string | undefined;
  /** The supplied face chosen for `font`. */
  face(font: PngRunFont): PngFontFace | undefined;
  /** Removes the export's faces from the font set. */
  release(): void;
}

export interface InstallPngFacesOptions {
  /** Prefix of the private family names, unique per format. */
  prefix: string;
  /** Every family and style the export paints. */
  fonts: Iterable<PngRunFont>;
  /** Faces supplied for this export. */
  faces: ReadonlyArray<PngFontFace>;
  /** The supplied face that paints `font`, chosen as the format's layout chooses. */
  choose(faces: ReadonlyArray<PngFontFace>, font: PngRunFont): PngFontFace | undefined;
  /** The canvas font string for `font` painted with `family`. */
  canvasFont(font: PngRunFont, family: string): string;
  /** What is exported, for error messages: "page" or "sheet". */
  subject: string;
}

const counters = new Map<string, number>();

export const pngFontKey = (font: PngRunFont): string => JSON.stringify([font.family, font.bold, font.italic, font.id ?? null]);

export function pngFontSet(): FontFaceSet | undefined {
  return typeof document !== 'undefined' ? document.fonts : (globalThis as { fonts?: FontFaceSet }).fonts;
}

export function fontUnavailable(family: string, subject: string): PngExportError {
  return new PngExportError('font-unavailable', `The font "${family}" could not load, so the ${subject} was not exported with a substitute.`);
}

/**
 * Installs the supplied faces an export uses under private family names, one
 * per face, so neither the editor's faces nor another export's can change what
 * it paints. Waits for every font the export paints, supplied or not, and
 * rejects with `font-unavailable` rather than paint a substitute.
 */
export async function installPngFaces(options: InstallPngFacesOptions): Promise<InstalledPngFaces> {
  const fonts = new Map<string, PngRunFont>();
  for (const font of options.fonts) fonts.set(pngFontKey(font), font);
  const set = pngFontSet();
  const installed = new Map<PngFontFace, string>();
  const chosen = new Map<string, { family: string; face: PngFontFace }>();
  const added: FontFace[] = [];
  const release = () => { for (const face of added) set?.delete(face); };
  const id = (counters.get(options.prefix) ?? 0) + 1;
  counters.set(options.prefix, id);
  try {
    for (const [key, font] of fonts) {
      const face = options.choose(options.faces, font);
      if (!face) continue;
      let family = installed.get(face);
      if (!family) {
        if (!set || typeof FontFace === 'undefined') throw fontUnavailable(face.family, options.subject);
        family = `${options.prefix}-${id}-${installed.size}`;
        const loaded = await new FontFace(family, face.bytes.slice().buffer as ArrayBuffer, {
          weight: (face.face?.bold ?? face.bold) ? '700' : '400', style: (face.face?.italic ?? face.italic) ? 'italic' : 'normal',
        }).load().catch(() => { throw fontUnavailable(face.family, options.subject); });
        set.add(loaded);
        added.push(loaded);
        installed.set(face, family);
      }
      chosen.set(key, { family, face });
    }
    if (set && typeof set.load === 'function') {
      for (const [key, font] of fonts) {
        const supplied = chosen.get(key);
        const matched = await set.load(options.canvasFont(font, supplied?.family ?? font.family))
          .catch(() => { throw fontUnavailable(font.family, options.subject); });
        if (supplied && !matched?.length) throw fontUnavailable(font.family, options.subject);
      }
    }
    return {
      family: (font) => chosen.get(pngFontKey(font))?.family,
      face: (font) => chosen.get(pngFontKey(font))?.face,
      release,
    };
  } catch (error) {
    release();
    throw error;
  }
}

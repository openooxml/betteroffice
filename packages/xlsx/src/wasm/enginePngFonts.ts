import type { BundledFontFace, BundledFontScript } from '@betteroffice/fonts';
import { isImplicitFontCdnError, reportImplicitCdnFailure } from '../../../../shared/fontCdnDiagnostics';
import { importFonts } from '../../../../shared/office-session/fontModules';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { pngFontKey, type PngFontFace, type PngRunFont } from '../../../../shared/office-session/pngFonts';
import type { SheetFontRequest } from '../render/png';
import { decodeFontFace } from './loader';

type Fonts = typeof import('@betteroffice/fonts');

let fontsModule: Promise<Fonts> | undefined;
/** The font packages, loaded with the first export that needs them; a worker resolves them only this way. */
const loadFonts = (): Promise<Fonts> => (fontsModule ??= importFonts().catch((error: unknown) => {
  fontsModule = undefined;
  throw error;
}));

type InstalledFace = (face: BundledFontFace | undefined) => Promise<BundledFontFace | undefined>;

/**
 * The face the native engine draws a cell font's Latin text with: the first installed face its font manager
 * resolves the family to, in its order (the family's bundled face, its glyph fallback, then Word's last resort).
 */
async function firstFace(fonts: Fonts, installedFace: InstalledFace, font: PngRunFont): Promise<BundledFontFace> {
  const primary = await installedFace(fonts.resolveBundledFamilyFace(font.family, font.bold, font.italic));
  if (primary) return primary;
  const glyph = await installedFace(fonts.resolveGlyphFallbackFace(font.family, font.bold, font.italic, 'word'));
  if (glyph) return glyph;
  const last = await installedFace(fonts.resolveLastResortChoice(font.family, font.bold, font.italic, 'word').face);
  return last ?? fonts.BASE_LATIN_FONTS.find((face) => (face.weight === 700) === font.bold && (face.style === 'italic') === font.italic)!;
}

const sameFace = (a: BundledFontFace, b: BundledFontFace) => a.family === b.family && a.weight === b.weight && a.style === b.style;

/**
 * The installed faces the native engine resolves one request to, in its font manager's order: the family's face,
 * a face per script of the text, the family's glyph fallback, then Word's last resort where the family has no face.
 */
async function chainFaces(fonts: Fonts, installedFace: InstalledFace, { font, scripts }: SheetFontRequest): Promise<BundledFontFace[]> {
  const faces: BundledFontFace[] = [];
  const push = (face: BundledFontFace | undefined) => { if (face && !faces.some((known) => sameFace(known, face))) faces.push(face); };
  const primary = await installedFace(fonts.resolveBundledFamilyFace(font.family, font.bold, font.italic));
  push(primary);
  for (const script of scripts) {
    push(await installedFace(fonts.resolveScriptFallbackFace(script as BundledFontScript, font.bold, font.italic, font.family)));
  }
  push(await installedFace(fonts.resolveGlyphFallbackFace(font.family, font.bold, font.italic, 'word')));
  const last = fonts.resolveLastResortChoice(font.family, font.bold, font.italic, 'word');
  if (!last.onlyIfPrimaryFails || !primary) push(await installedFace(last.face));
  return faces;
}

/**
 * The faces the native engine draws each font of a sheet with, in order. A font whose text is all in scripts the
 * bundled Latin faces cover gets the one face that draws it, so such a sheet loads nothing more. A font set in
 * another script anywhere on the sheet gets the native chain: every request for that font resolved in
 * display-list order, the faces joined without repeats, so a character falls through them as it does natively.
 */
async function sheetChains(fonts: Fonts, installedFace: InstalledFace, requests: readonly SheetFontRequest[]): Promise<Map<string, { font: PngRunFont; faces: BundledFontFace[] }>> {
  const scripted = new Set(requests.filter((request) => request.scripts.length).map((request) => pngFontKey(request.font)));
  const chains = new Map<string, { font: PngRunFont; faces: BundledFontFace[] }>();
  for (const request of requests) {
    const key = pngFontKey(request.font);
    const chain = chains.get(key) ?? { font: request.font, faces: [] };
    chains.set(key, chain);
    if (!scripted.has(key)) {
      if (!chain.faces.length) chain.faces.push(await firstFace(fonts, installedFace, request.font));
      continue;
    }
    for (const face of await chainFaces(fonts, installedFace, request)) if (!chain.faces.some((known) => sameFace(known, face))) chain.faces.push(face);
  }
  for (const chain of chains.values()) if (!chain.faces.length) chain.faces.push(await firstFace(fonts, installedFace, chain.font));
  return chains;
}

/** The faces a canvas export paints a sheet with when no `fonts` are passed, and each font's fallbacks in order. */
export interface DefaultSheetFaces {
  faces: PngFontFace[];
  fallbacks(font: PngRunFont): PngFontFace[];
}

/**
 * Bundled faces for the fonts a sheet's text paints with, as the native engine draws them, so a canvas export
 * without `fonts` matches it. Each face serves the style the text asks for and installs with its own weight and
 * style, so bold Calibri Light drawn with a regular face gets synthetic bold, and regular Arial Black its bold face,
 * as natively. An abort rejects at once without waiting for loads other exports may share.
 */
export async function defaultSheetFaces(
  requests: readonly SheetFontRequest[], signal?: AbortSignal,
): Promise<DefaultSheetFaces> {
  signal?.throwIfAborted();
  const loading = loadFonts().then(async (manager) => {
    const unavailable = new Set<string>();
    const reported = new Set<Error>();
    const installedFace: InstalledFace = async (face) =>
      face && !unavailable.has(face.file) && await manager.bundledFaceInstalled(face).catch(() => true) ? face : undefined;
    const loadFace = async (face: BundledFontFace, family: string): Promise<PngFontFace | undefined> => {
      try {
        const bytes = new Uint8Array(await manager.loadBundledFontBytes(face, undefined, { decoder: decodeFontFace, strict: true }));
        const own = { bold: face.weight === 700, italic: face.style === 'italic' };
        return { family: face.family, ...own, bytes, face: own };
      } catch (error) {
        if (isImplicitFontCdnError(error) && (error.diagnostic.pack === 'cjk' || error.diagnostic.pack === 'intl')) {
          unavailable.add(face.file);
          if (!reported.has(error)) {
            reported.add(error);
            reportImplicitCdnFailure(error);
          }
          return undefined;
        }
        const reason = error instanceof Error ? error.message : String(error);
        throw new PngExportError('font-unavailable', `${face.family} for ${family} could not load: ${reason}`);
      }
    };
    const shared = new Map<string, Promise<PngFontFace | undefined>>();
    const sharedFace = (face: BundledFontFace, family: string) => {
      let pending = shared.get(face.file);
      if (!pending) {
        pending = loadFace(face, family);
        shared.set(face.file, pending);
      }
      return pending;
    };
    for (;;) {
      const chains = await sheetChains(manager, installedFace, requests);
      const excluded = unavailable.size;
      const faces: PngFontFace[] = [];
      const fallbacks = new Map<string, PngFontFace[]>();
      await Promise.all([...chains].map(async ([key, { font, faces: chain }]) => {
        const [primary, ...others] = await Promise.all(chain.map((face) => sharedFace(face, font.family)));
        if (primary) faces.push({ ...primary, family: font.family, bold: font.bold, italic: font.italic });
        const available = others.filter((face) => face !== undefined);
        if (available.length) fallbacks.set(key, available);
      }));
      if (unavailable.size === excluded) {
        return { faces, fallbacks: (font: PngRunFont) => [...(fallbacks.get(pngFontKey(font)) ?? [])] };
      }
      signal?.throwIfAborted();
    }
  });
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

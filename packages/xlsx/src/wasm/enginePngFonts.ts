import type { BundledFontFace, BundledFontScript } from '@betteroffice/fonts';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { pngFontKey, type PngFontFace, type PngRunFont } from '../../../../shared/office-session/pngFonts';
import type { SheetFontRequest } from '../render/png';
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

const sameFace = (a: BundledFontFace, b: BundledFontFace) => a.family === b.family && a.weight === b.weight && a.style === b.style;

/**
 * The installed faces the native engine resolves one request to, in its font manager's order: the family's face,
 * a face per script of the text, the family's glyph fallback, then Word's last resort where the family has no face.
 */
async function chainFaces(fonts: Fonts, { font, scripts }: SheetFontRequest): Promise<BundledFontFace[]> {
  const faces: BundledFontFace[] = [];
  const push = (face: BundledFontFace | undefined) => { if (face && !faces.some((known) => sameFace(known, face))) faces.push(face); };
  const primary = await installedFace(fonts, fonts.resolveBundledFamilyFace(font.family, font.bold, font.italic));
  push(primary);
  for (const script of scripts) {
    push(await installedFace(fonts, fonts.resolveScriptFallbackFace(script as BundledFontScript, font.bold, font.italic)));
  }
  push(await installedFace(fonts, fonts.resolveGlyphFallbackFace(font.family, font.bold, font.italic, 'word')));
  const last = fonts.resolveLastResortChoice(font.family, font.bold, font.italic, 'word');
  if (!last.onlyIfPrimaryFails || !primary) push(await installedFace(fonts, last.face));
  return faces;
}

/**
 * The faces the native engine draws each font of a sheet with, in order. A font whose text is all in scripts the
 * bundled Latin faces cover gets the one face that draws it, so such a sheet loads nothing more. A font set in
 * another script anywhere on the sheet gets the native chain: every request for that font resolved in
 * display-list order, the faces joined without repeats, so a character falls through them as it does natively.
 */
async function sheetChains(fonts: Fonts, requests: readonly SheetFontRequest[]): Promise<Map<string, { font: PngRunFont; faces: BundledFontFace[] }>> {
  const scripted = new Set(requests.filter((request) => request.scripts.length).map((request) => pngFontKey(request.font)));
  const chains = new Map<string, { font: PngRunFont; faces: BundledFontFace[] }>();
  for (const request of requests) {
    const key = pngFontKey(request.font);
    const chain = chains.get(key) ?? { font: request.font, faces: [] };
    chains.set(key, chain);
    if (!scripted.has(key)) {
      if (!chain.faces.length) chain.faces.push(await firstFace(fonts, request.font));
      continue;
    }
    for (const face of await chainFaces(fonts, request)) if (!chain.faces.some((known) => sameFace(known, face))) chain.faces.push(face);
  }
  for (const chain of chains.values()) if (!chain.faces.length) chain.faces.push(await firstFace(fonts, chain.font));
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
 * as natively. Faces load only from installed font packages, and only for the fonts that need them. A face that
 * cannot load rejects with `font-unavailable` naming it; an abort rejects at once without waiting for loads other
 * exports may share.
 */
export async function defaultSheetFaces(requests: readonly SheetFontRequest[], signal?: AbortSignal): Promise<DefaultSheetFaces> {
  signal?.throwIfAborted();
  const loading = loadFonts().then(async (manager) => {
    const chains = await sheetChains(manager, requests);
    const bytes = (face: BundledFontFace, family: string) => load(manager, face).catch((error: unknown) => {
      const reason = error instanceof Error ? error.message : String(error);
      throw new PngExportError('font-unavailable', `${face.family} for ${family} could not load: ${reason}`);
    });
    const own = (face: BundledFontFace) => ({ bold: face.weight === 700, italic: face.style === 'italic' });
    const shared = new Map<string, Promise<PngFontFace>>();
    const fallbackFace = (face: BundledFontFace, family: string) => {
      let pending = shared.get(face.file);
      if (!pending) {
        pending = bytes(face, family).then((data) => ({ family: face.family, ...own(face), bytes: data, face: own(face) }));
        shared.set(face.file, pending);
      }
      return pending;
    };
    const faces: PngFontFace[] = [];
    const fallbacks = new Map<string, PngFontFace[]>();
    await Promise.all([...chains].map(async ([key, { font, faces: chain }]) => {
      const [first, ...rest] = chain;
      const [primary, ...others] = await Promise.all([
        bytes(first!, font.family).then((data) => ({ family: font.family, bold: font.bold, italic: font.italic, bytes: data, face: own(first!) })),
        ...rest.map((face) => fallbackFace(face, font.family)),
      ]);
      faces.push(primary!);
      if (others.length) fallbacks.set(key, others);
    }));
    return { faces, fallbacks: (font: PngRunFont) => [...(fallbacks.get(pngFontKey(font)) ?? [])] };
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

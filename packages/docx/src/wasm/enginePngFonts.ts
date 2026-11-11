import type { BundledFontFace, BundledFontScript } from '@betteroffice/fonts';
import { fontUnavailable } from '../layout/render/png';
import { decodeFontFace } from './fontDecoder';

type Fonts = typeof import('@betteroffice/fonts');

/** A font a document lays text out with, and the scripts of that text. */
export interface DocumentFontRequest {
  family: string;
  bold: boolean;
  italic: boolean;
  scripts?: readonly BundledFontScript[];
}

let fontsModule: Promise<Fonts> | undefined;
/** The font packages, loaded with the first export that needs them. */
const loadFonts = (): Promise<Fonts> => (fontsModule ??= import('@betteroffice/fonts').catch((error: unknown) => {
  fontsModule = undefined;
  throw error;
}));

/** An installed face, or undefined for one whose package is not installed. */
async function installedFace(fonts: Fonts, face: BundledFontFace | undefined): Promise<BundledFontFace | undefined> {
  return face && await fonts.bundledFaceInstalled(face).catch(() => true) ? face : undefined;
}

const sameFace = (a: BundledFontFace, b: BundledFontFace) => a.family === b.family && a.weight === b.weight && a.style === b.style;

/**
 * The installed faces the native engine resolves one font to, in its font manager's order: the family's face,
 * a face per script of the text, the family's glyph fallback, then Word's last resort where the family has no face.
 */
export async function documentFaceChain(font: DocumentFontRequest): Promise<BundledFontFace[]> {
  const fonts = await loadFonts();
  const faces: BundledFontFace[] = [];
  const push = (face: BundledFontFace | undefined) => {
    if (face && !faces.some((known) => sameFace(known, face))) faces.push(face);
  };
  const primary = await installedFace(fonts, fonts.resolveBundledFamilyFace(font.family, font.bold, font.italic));
  push(primary);
  for (const script of font.scripts ?? []) {
    push(await installedFace(fonts, fonts.resolveScriptFallbackFace(script, font.bold, font.italic, font.family)));
  }
  push(await installedFace(fonts, fonts.resolveGlyphFallbackFace(font.family, font.bold, font.italic, 'word')));
  const last = fonts.resolveLastResortChoice(font.family, font.bold, font.italic, 'word');
  if (!last.onlyIfPrimaryFails || !primary) push(await installedFace(fonts, last.face));
  return faces;
}

/** A face's bytes, loaded strictly from its package; one that cannot load rejects with `font-unavailable` naming it. */
export async function loadDocumentFace(face: BundledFontFace, family: string): Promise<Uint8Array> {
  const fonts = await loadFonts();
  try {
    return new Uint8Array(await fonts.loadBundledFontBytes(face, undefined, { decoder: decodeFontFace, strict: true }));
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw fontUnavailable(`${face.family} for ${family} could not load: ${reason}`);
  }
}

/**
 * Loads every installed face the native engine draws the given fonts with, each once, so an export without a
 * provider of its own rejects with `font-unavailable` for a face it cannot load instead of drawing without it.
 * Only the faces those fonts resolve to load; an abort rejects at once.
 */
export async function loadDocumentFaces(requests: readonly DocumentFontRequest[], signal?: AbortSignal): Promise<void> {
  signal?.throwIfAborted();
  const loading = (async () => {
    const families = new Map<string, { face: BundledFontFace; family: string }>();
    for (const request of requests) {
      for (const face of await documentFaceChain(request)) {
        if (!families.has(face.file)) families.set(face.file, { face, family: request.family });
      }
    }
    await Promise.all([...families.values()].map(({ face, family }) => loadDocumentFace(face, family)));
  })();
  return untilAborted(loading, signal);
}

/** `work`'s result, or a rejection with the signal's reason as soon as it aborts; `work` itself runs on. */
export async function untilAborted<T>(work: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return work;
  signal.throwIfAborted();
  let stop: (() => void) | undefined;
  const aborted = new Promise<never>((_, reject) => {
    stop = () => reject(signal.reason);
    signal.addEventListener('abort', stop, { once: true });
  });
  work.catch(() => {});
  try {
    return await Promise.race([work, aborted]);
  } finally {
    signal.removeEventListener('abort', stop!);
  }
}

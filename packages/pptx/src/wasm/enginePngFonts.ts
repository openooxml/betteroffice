import type { BundledFontFace, BundledFontScript } from '@betteroffice/fonts';
import { isImplicitFontCdnError, reportImplicitCdnFailure } from '../../../../shared/fontCdnDiagnostics';
import { pngFontCdnCacheKey } from '../../../../shared/office-session/fontCdnRuntime';
import { importFonts } from '../../../../shared/office-session/fontModules';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { RequestedFontFace } from '../../../../shared/pptxDefaultFonts';
import type { PptxFontFace, SlideDisplayList } from '../types';
import type { PptxDocument } from './generated/pptx_wasm.js';
import { PptxRenderer } from './engine';
import { decodeFontFace } from './loader';

type ExportFace = PptxFontFace & { id: number; fallback?: boolean };
type Fonts = typeof import('@betteroffice/fonts');
type InstalledFace = (face: BundledFontFace | undefined) => Promise<BundledFontFace | undefined>;

let fontsModule: Promise<Fonts> | undefined;
/** The font packages, loaded with the first export that needs them; a worker resolves them only this way. */
const loadFonts = (): Promise<Fonts> => (fontsModule ??= importFonts().catch((error: unknown) => {
  fontsModule = undefined;
  throw error;
}));

interface Exporter {
  renderer: PptxRenderer;
  fontCdn: string;
  measured: ExportFace[];
  sources: Uint8Array[];
  faces: ExportFace[];
  owned: Set<string>;
  primary: Set<string>;
  fallback: Set<string>;
  requested: Set<string>;
  queue: Promise<unknown>;
  users: number;
  released: boolean;
}

const exporters = new WeakMap<object, Exporter>();
/** Loading rounds an export waits through while edits keep asking for new faces. */
const MAX_FONT_ROUNDS = 4;

const familyKey = (family: string) => family.trim().toLowerCase();
const faceKey = (family: string, bold: boolean, italic: boolean) => JSON.stringify([familyKey(family), bold, italic]);
const isBold = (face: BundledFontFace) => face.weight === 700;
const isItalic = (face: BundledFontFace) => face.style === 'italic';

function register(exporter: Exporter, face: PptxFontFace, fallback: boolean): void {
  const bold = face.bold ?? false;
  const italic = face.italic ?? false;
  const id = fallback
    ? exporter.renderer.registerFallbackFont(face.family, bold, italic, face.bytes)
    : exporter.renderer.registerFont(face.family, bold, italic, face.bytes);
  exporter.faces.push({ id, family: face.family, bold, italic, bytes: face.bytes, fallback });
  (fallback ? exporter.fallback : exporter.primary).add(faceKey(face.family, bold, italic));
}

function create(measured: ExportFace[], added: ReadonlyArray<PptxFontFace>, fontCdn: string): Exporter {
  const exporter: Exporter = {
    renderer: new PptxRenderer(), fontCdn, measured, sources: measured.map((face) => face.bytes), faces: [],
    owned: new Set(), primary: new Set(), fallback: new Set(), requested: new Set(),
    queue: Promise.resolve(), users: 0, released: false,
  };
  try {
    for (const face of [...measured, ...added]) {
      const fallback = 'fallback' in face && face.fallback === true;
      register(exporter, face, fallback);
      if (!fallback) exporter.owned.add(familyKey(face.family));
    }
  } catch (error) {
    exporter.renderer.free();
    throw error;
  }
  return exporter;
}

/**
 * The installed bundled faces one request resolves to, in the order the native font manager gives them; like it,
 * faces of packages that are not installed are left out before the primary and the last resort are chosen.
 */
async function resolve(fonts: Fonts, installedFace: InstalledFace, request: RequestedFontFace): Promise<BundledFontFace[]> {
  const faces: BundledFontFace[] = [];
  const push = (face: BundledFontFace | undefined) => {
    if (face && !faces.some((known) => known.family === face.family && isBold(known) === isBold(face) && isItalic(known) === isItalic(face))) {
      faces.push(face);
    }
  };
  const primary = await installedFace(fonts.resolveBundledFamilyFace(request.family, request.bold, request.italic));
  push(primary);
  for (const script of request.scripts ?? []) {
    push(await installedFace(fonts.resolveScriptFallbackFace(script as BundledFontScript, request.bold, request.italic, request.family)));
  }
  push(await installedFace(fonts.resolveGlyphFallbackFace(request.family, request.bold, request.italic, 'powerpoint')));
  const last = fonts.resolveLastResortChoice(request.family, request.bold, request.italic, 'powerpoint');
  if (!last.onlyIfPrimaryFails || !primary) push(await installedFace(last.face));
  if (!faces.length) push(fonts.BASE_LATIN_FONTS.find((face) => isBold(face) === request.bold && isItalic(face) === request.italic));
  return faces;
}

const requestKey = (request: RequestedFontFace) =>
  JSON.stringify([familyKey(request.family), request.bold, request.italic, request.scripts ?? []]);

const hasPrimary = (exporter: Exporter, request: RequestedFontFace) =>
  exporter.owned.has(familyKey(request.family)) || exporter.primary.has(faceKey(request.family, request.bold, request.italic));

const hasFallback = (exporter: Exporter, face: BundledFontFace) => exporter.fallback.has(faceKey(face.family, isBold(face), isItalic(face)));

/** The requests of the slide's current version the export renderer has not served yet. */
function missingRequests(exporter: Exporter, document: PptxDocument, slideId: string, suppliedFamilies: ReadonlySet<string>): RequestedFontFace[] {
  const { faces: requests } = JSON.parse(exporter.renderer.requestedFontFacesJson(document, slideId)) as { faces: RequestedFontFace[] };
  return requests.filter((request) => !suppliedFamilies.has(familyKey(request.family)) && !exporter.requested.has(requestKey(request)));
}

async function loadDefaults(
  exporter: Exporter, wanted: RequestedFontFace[], unavailable: Set<string>, reported: Set<Error>, signal?: AbortSignal,
) {
  const fonts = await loadFonts();
  const installedFace: InstalledFace = async (face) =>
    face && !unavailable.has(face.file) && await fonts.bundledFaceInstalled(face).catch(() => true) ? face : undefined;
  const load = async (face: BundledFontFace, family: string): Promise<Uint8Array | undefined> => {
    try {
      return new Uint8Array(await fonts.loadBundledFontBytes(face, undefined, { decoder: decodeFontFace, strict: true }));
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
  const pending = new Map<string, Promise<Uint8Array | undefined>>();
  const loadFace = (face: BundledFontFace, family: string) => {
    let bytes = pending.get(face.file);
    if (!bytes) {
      bytes = load(face, family);
      pending.set(face.file, bytes);
    }
    return bytes;
  };
  for (;;) {
    const plans = await Promise.all(wanted.map(async (request) => ({ request, faces: await resolve(fonts, installedFace, request) })));
    const excluded = unavailable.size;
    const loaded = await Promise.all(plans.map(async ({ request, faces }) => {
      const primary = hasPrimary(exporter, request);
      const bytes = await Promise.all(faces.map((face, index) => (index === 0 && !primary) || !hasFallback(exporter, face)
        ? loadFace(face, request.family) : undefined));
      return { request, faces, bytes };
    }));
    signal?.throwIfAborted();
    if (unavailable.size === excluded) return loaded;
  }
}

function registerDefaults(exporter: Exporter, plans: Awaited<ReturnType<typeof loadDefaults>>): void {
  for (const { request, faces, bytes } of plans) {
    if (exporter.requested.has(requestKey(request))) continue;
    const primary = hasPrimary(exporter, request);
    for (const [index, face] of faces.entries()) {
      const data = bytes[index];
      if (index === 0 && !primary) {
        if (data) register(exporter, { family: request.family, bold: request.bold, italic: request.italic, bytes: data }, false);
      } else if (!hasFallback(exporter, face)) {
        if (data) register(exporter, { family: face.family, bold: isBold(face), italic: isItalic(face), bytes: data }, true);
      }
    }
    exporter.requested.add(requestKey(request));
  }
}

const current = (exporter: Exporter, measured: ExportFace[]) => exporter.measured.length === measured.length &&
  measured.every((face, index) => exporter.measured[index] === face && exporter.sources[index] === face.bytes);

function slideIndex(document: PptxDocument, slideId: string): number {
  const { slides } = JSON.parse(document.snapshotJson()) as { slides: Array<{ id: string }> };
  const index = slides.findIndex((candidate) => candidate.id === slideId);
  if (index < 0) throw new RangeError(`no slide with ID ${JSON.stringify(slideId)}`);
  return index;
}

function release(exporter: Exporter): void {
  exporter.released = true;
  if (exporter.users === 0) exporter.renderer.free();
}

function exportFontCdnKey(signal?: AbortSignal): Promise<string> {
  if (!signal) return pngFontCdnCacheKey();
  signal.throwIfAborted();
  let abort!: () => void;
  const cancelled = new Promise<never>((_, reject) => { abort = () => reject(signal.reason); });
  signal.addEventListener('abort', abort, { once: true });
  return Promise.race([pngFontCdnCacheKey(), cancelled]).finally(() => signal.removeEventListener('abort', abort));
}

/**
 * Lays the slide with `slideId` out again in an export renderer holding the faces the presentation registered, any
 * new family in `added`, and the bundled faces the slide asks for, as the native export does, so both measure alike.
 * Faces are collected again until the document stays unchanged while they load, at most `MAX_FONT_ROUNDS` times;
 * the layout then runs without yielding, on that version and at the slide's index in it. Edits that keep asking
 * for new faces past that reject with `font-unavailable`, so the caller can export again.
 */
export async function defaultFontLayout(
  document: PptxDocument, slideId: string, measured: ExportFace[], added: ReadonlyArray<PptxFontFace> | undefined, signal?: AbortSignal,
): Promise<{ list: SlideDisplayList; fonts: ExportFace[] }> {
  slideIndex(document, slideId);
  const known = new Set(measured.map((face) => familyKey(face.family)));
  const supplied = added?.filter((face) => !known.has(familyKey(face.family))) ?? [];
  const suppliedFamilies = new Set(supplied.map((face) => familyKey(face.family)));
  const fontCdn = await exportFontCdnKey(signal);
  signal?.throwIfAborted();
  let exporter = supplied.length ? undefined : exporters.get(document);
  if (exporter && (exporter.fontCdn !== fontCdn || !current(exporter, measured))) {
    releaseDefaultFontLayout(document);
    exporter = undefined;
  }
  if (!exporter) {
    exporter = create(measured, supplied, fontCdn);
    if (supplied.length) exporter.released = true;
    else exporters.set(document, exporter);
  }
  const held = exporter;
  held.users += 1;
  const run = held.queue.then(async () => {
    const unavailable = new Set<string>();
    const reported = new Set<Error>();
    let layout = held;
    try {
      for (let round = 0; ; round += 1) {
        const version = document.documentVersion();
        slideIndex(document, slideId);
        const wanted = missingRequests(layout, document, slideId, suppliedFamilies);
        if (!wanted.length) break;
        if (round === MAX_FONT_ROUNDS) {
          throw new PngExportError('font-unavailable', 'the faces this slide needs kept changing while they loaded; export again');
        }
        const plans = await loadDefaults(layout, wanted, unavailable, reported, signal);
        if (unavailable.size && !supplied.length && layout === held) {
          layout = create([], held.faces, fontCdn);
          layout.owned = new Set(held.owned);
          layout.requested = new Set(held.requested);
        }
        registerDefaults(layout, plans);
        if (document.documentVersion() === version) break;
      }
      signal?.throwIfAborted();
      const list = JSON.parse(layout.renderer.exportSlideJson(document, slideIndex(document, slideId))) as SlideDisplayList;
      return { list, fonts: [...layout.faces] };
    } finally {
      if (layout !== held) layout.renderer.free();
    }
  });
  held.queue = run.catch(() => {});
  try {
    return await run;
  } finally {
    held.users -= 1;
    if (held.released && held.users === 0) held.renderer.free();
  }
}

/** Frees the export renderer kept for `document` once no export uses it. */
export function releaseDefaultFontLayout(document: object): void {
  const exporter = exporters.get(document);
  exporters.delete(document);
  if (exporter) release(exporter);
}

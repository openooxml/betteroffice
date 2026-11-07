import {
  bundledFaceInstalled, loadBundledFontBytes, resolveBundledFamilyFace, resolveGlyphFallbackFace, resolveLastResortChoice, resolveScriptFallbackFace,
  type BundledFontFace, type BundledFontScript,
} from '@betteroffice/fonts';
import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { RequestedFontFace } from '../../../../shared/pptxDefaultFonts';
import type { PptxFontFace, SlideDisplayList } from '../types';
import type { PptxDocument } from './generated/pptx_wasm.js';
import { PptxRenderer } from './engine';
import { decodeFontFace } from './loader';

type ExportFace = PptxFontFace & { id: number; fallback?: boolean };

interface Exporter {
  renderer: PptxRenderer;
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
const loaded = new Map<string, Promise<Uint8Array>>();

const familyKey = (family: string) => family.trim().toLowerCase();
const faceKey = (family: string, bold: boolean, italic: boolean) => JSON.stringify([familyKey(family), bold, italic]);
const isBold = (face: BundledFontFace) => face.weight === 700;
const isItalic = (face: BundledFontFace) => face.style === 'italic';

function load(face: BundledFontFace): Promise<Uint8Array> {
  let pending = loaded.get(face.file);
  if (!pending) {
    const started = loadBundledFontBytes(face, undefined, { decoder: decodeFontFace, strict: true }).then((buffer) => new Uint8Array(buffer));
    started.catch(() => { if (loaded.get(face.file) === started) loaded.delete(face.file); });
    loaded.set(face.file, started);
    pending = started;
  }
  return pending;
}

/** A face the slide needs, or the error the native export gives when it cannot load. */
async function needed(face: BundledFontFace, family: string): Promise<Uint8Array> {
  try {
    return await load(face);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new PngExportError('font-unavailable', `${face.family} for ${family} could not load: ${reason}`);
  }
}

function register(exporter: Exporter, face: PptxFontFace, fallback: boolean): void {
  const bold = face.bold ?? false;
  const italic = face.italic ?? false;
  const id = fallback
    ? exporter.renderer.registerFallbackFont(face.family, bold, italic, face.bytes)
    : exporter.renderer.registerFont(face.family, bold, italic, face.bytes);
  exporter.faces.push({ id, family: face.family, bold, italic, bytes: face.bytes, fallback });
  (fallback ? exporter.fallback : exporter.primary).add(faceKey(face.family, bold, italic));
}

function create(measured: ExportFace[], added: ReadonlyArray<PptxFontFace>): Exporter {
  const exporter: Exporter = {
    renderer: new PptxRenderer(), measured, sources: measured.map((face) => face.bytes), faces: [],
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

/** An installed face, or undefined for one whose package is not installed. */
async function installedFace(face: BundledFontFace | undefined): Promise<BundledFontFace | undefined> {
  return face && await bundledFaceInstalled(face).catch(() => true) ? face : undefined;
}

/**
 * The installed bundled faces one request resolves to, in the order the native font manager gives them; like it,
 * faces of packages that are not installed are left out before the primary and the last resort are chosen.
 */
async function resolve(request: RequestedFontFace): Promise<BundledFontFace[]> {
  const faces: BundledFontFace[] = [];
  const push = (face: BundledFontFace | undefined) => {
    if (face && !faces.some((known) => known.family === face.family && isBold(known) === isBold(face) && isItalic(known) === isItalic(face))) {
      faces.push(face);
    }
  };
  const primary = await installedFace(resolveBundledFamilyFace(request.family, request.bold, request.italic));
  push(primary);
  for (const script of request.scripts ?? []) {
    push(await installedFace(resolveScriptFallbackFace(script as BundledFontScript, request.bold, request.italic)));
  }
  push(await installedFace(resolveGlyphFallbackFace(request.family, request.bold, request.italic, 'powerpoint')));
  const last = resolveLastResortChoice(request.family, request.bold, request.italic, 'powerpoint');
  if (!last.onlyIfPrimaryFails || !primary) push(await installedFace(last.face));
  return faces;
}

const requestKey = (request: RequestedFontFace) =>
  JSON.stringify([familyKey(request.family), request.bold, request.italic, request.scripts ?? []]);

const hasPrimary = (exporter: Exporter, request: RequestedFontFace) =>
  exporter.owned.has(familyKey(request.family)) || exporter.primary.has(faceKey(request.family, request.bold, request.italic));

const hasFallback = (exporter: Exporter, face: BundledFontFace) => exporter.fallback.has(faceKey(face.family, isBold(face), isItalic(face)));

/** The requests of the slide's current version the export renderer has not served yet. */
function missingRequests(exporter: Exporter, document: PptxDocument, slideId: string): RequestedFontFace[] {
  const { faces: requests } = JSON.parse(exporter.renderer.requestedFontFacesJson(document, slideId)) as { faces: RequestedFontFace[] };
  return requests.filter((request) => !exporter.requested.has(requestKey(request)));
}

/**
 * Loads the faces `wanted` resolves to, then registers them in the native engine's order: the first installed face
 * as the family's primary unless the family has one, the rest as fallbacks. Nothing registers until every load
 * has settled, so a failed pass leaves the renderer as it was.
 */
async function registerDefaults(exporter: Exporter, wanted: RequestedFontFace[], signal?: AbortSignal): Promise<void> {
  const plans = await Promise.all(wanted.map(async (request) => {
    const faces = await resolve(request);
    const primary = hasPrimary(exporter, request);
    const loads = faces.map((face, index) => (index === 0 && !primary) || !hasFallback(exporter, face)
      ? needed(face, request.family) : undefined);
    return { request, faces, loads };
  }));
  const bytes = await Promise.all(plans.map((plan) => Promise.all(plan.loads)));
  signal?.throwIfAborted();
  for (const [planIndex, { request, faces }] of plans.entries()) {
    if (exporter.requested.has(requestKey(request))) continue;
    const primary = hasPrimary(exporter, request);
    for (const [index, face] of faces.entries()) {
      const data = bytes[planIndex]![index];
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
  let exporter = supplied.length ? undefined : exporters.get(document);
  if (exporter && !current(exporter, measured)) {
    releaseDefaultFontLayout(document);
    exporter = undefined;
  }
  if (!exporter) {
    exporter = create(measured, supplied);
    if (supplied.length) exporter.released = true;
    else exporters.set(document, exporter);
  }
  const held = exporter;
  held.users += 1;
  const run = held.queue.then(async () => {
    for (let round = 0; ; round += 1) {
      const version = document.documentVersion();
      const wanted = missingRequests(held, document, slideId);
      if (!wanted.length) break;
      if (round === MAX_FONT_ROUNDS) {
        throw new PngExportError('font-unavailable', 'the faces this slide needs kept changing while they loaded; export again');
      }
      await registerDefaults(held, wanted, signal);
      if (document.documentVersion() === version) break;
    }
    signal?.throwIfAborted();
    const list = JSON.parse(held.renderer.exportSlideJson(document, slideIndex(document, slideId))) as SlideDisplayList;
    return { list, fonts: [...held.faces] };
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

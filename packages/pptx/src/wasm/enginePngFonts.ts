import {
  loadBundledFontBytes, resolveBundledFamilyFace, resolveGlyphFallbackFace, resolveLastResortChoice, resolveScriptFallbackFace,
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

/**
 * A face the slide needs, the error the native export gives when it cannot load, or undefined when its add-on
 * package is not installed, which the native export leaves out of the slide's faces the same way.
 */
async function needed(face: BundledFontFace, family: string): Promise<Uint8Array | undefined> {
  try {
    return await load(face);
  } catch (error) {
    if (error instanceof Error && error.name === 'MissingFontAddonError') return undefined;
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

/** The bundled faces one request resolves to, in the order the native font manager gives them. */
function resolve(request: RequestedFontFace): BundledFontFace[] {
  const faces: BundledFontFace[] = [];
  const push = (face: BundledFontFace | undefined) => {
    if (face && !faces.some((known) => known.family === face.family && isBold(known) === isBold(face) && isItalic(known) === isItalic(face))) {
      faces.push(face);
    }
  };
  const primary = resolveBundledFamilyFace(request.family, request.bold, request.italic);
  push(primary);
  for (const script of request.scripts ?? []) push(resolveScriptFallbackFace(script as BundledFontScript, request.bold, request.italic));
  push(resolveGlyphFallbackFace(request.family, request.bold, request.italic, 'powerpoint'));
  const last = resolveLastResortChoice(request.family, request.bold, request.italic, 'powerpoint');
  if (!last.onlyIfPrimaryFails || !primary) push(last.face);
  return faces;
}

async function registerDefaults(exporter: Exporter, document: PptxDocument, slideId: string, signal?: AbortSignal): Promise<void> {
  const { faces: requests } = JSON.parse(exporter.renderer.requestedFontFacesJson(document, slideId)) as { faces: RequestedFontFace[] };
  const plans = requests.map((request) => ({ request, faces: resolve(request) }));
  for (const plan of plans) for (const face of plan.faces) load(face).catch(() => {});
  for (const { request, faces } of plans) {
    signal?.throwIfAborted();
    const key = JSON.stringify([familyKey(request.family), request.bold, request.italic, request.scripts ?? []]);
    if (exporter.requested.has(key)) continue;
    const hasPrimary = exporter.owned.has(familyKey(request.family)) ||
      exporter.primary.has(faceKey(request.family, request.bold, request.italic));
    for (const [index, face] of faces.entries()) {
      if (index === 0 && !hasPrimary) {
        const bytes = await needed(face, request.family);
        if (bytes) register(exporter, { family: request.family, bold: request.bold, italic: request.italic, bytes }, false);
      } else if (!exporter.fallback.has(faceKey(face.family, isBold(face), isItalic(face)))) {
        const bytes = await needed(face, request.family);
        if (bytes) register(exporter, { family: face.family, bold: isBold(face), italic: isItalic(face), bytes }, true);
      }
    }
    exporter.requested.add(key);
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
 * Lays the slide with `slideId` out again, at the index it has once the faces are in, in an export renderer holding
 * the faces the presentation registered, any new family in `added`, and the bundled faces the slide asks for, as the
 * native export does, so both measure alike.
 */
export async function defaultFontLayout(
  document: PptxDocument, slideId: string, measured: ExportFace[], added: ReadonlyArray<PptxFontFace> | undefined, signal?: AbortSignal,
): Promise<{ list: SlideDisplayList; fonts: ExportFace[] }> {
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
    slideIndex(document, slideId);
    await registerDefaults(held, document, slideId, signal);
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

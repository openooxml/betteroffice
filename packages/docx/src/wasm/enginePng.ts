import { GlyphCache } from '../layout/render/glyphCache';
import { createCanvasImageResolver } from '../layout/render/canvasImageResolver';
import { rasterizeDisplayPageToBackBuffer, type PageCanvasBuffer } from '../layout/render/canvasBackend';
import { canvasPngAvailable, canvasUnavailable, fontUnavailable, pagePngRequest, type DocxPngOptions } from '../layout/render/png';
import type { PngPictureUnavailableReason, PngSkippedEffectKind, PngUnavailablePicture } from '../../../../shared/office-session/pngExportError';
import { DEFAULT_STRICT_PICTURES, strictPicturesError } from '../../../../shared/office-session/strictPictures';
import type { DisplayList, DisplayPage, DisplayPrimitive } from '../layout/render/displayList';
import { skippedEffectKinds } from './engineEdit';

/** What an engine needs to draw one page, from the session that laid it out. */
export interface DocxPngEngine {
  displayList: DisplayList;
  /** The font ids layout measured each family with, for engines that rasterize the list themselves. */
  fontChains: Record<string, number[]>;
  page: DisplayPage;
  outlineGlyph(fontId: number, glyphId: number): string;
  media(token: string): { bytes: Uint8Array; mimeType: string } | null;
  /** The session, for engines that render PNG themselves. */
  session: unknown;
}

/** Which engine this module draws with; the native package swaps in its own. */
export const engineKind: 'wasm' | 'native' = 'wasm';

const SVG_DATA_URL = /^data:image\/svg/i;
const METAFILE_PLACEHOLDER = /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" width="[\d.]+" height="[\d.]+" viewBox="0 0 [\d.]+ [\d.]+" preserveAspectRatio="none"><rect x="0\.5" y="0\.5" width="[\d.]+" height="[\d.]+" fill="#f1f3f4" stroke="#c4c7c5"\/><g transform="translate\([\d.]+ [\d.]+\) scale\([\d.]+\)" fill="none" stroke="#9aa0a6" stroke-width="1\.5" stroke-linejoin="round"><rect x="-12" y="-9" width="24" height="18" rx="2"\/><circle cx="-5" cy="-3\.5" r="2"\/><path d="M-10 7L-3 0L2 5L5 2L10 7"\/><\/g><\/svg>$/;

/** The kind each name the canvas walk reports stands for. */
const CANVAS_EFFECT_KINDS: Record<PngSkippedEffectKind, PngSkippedEffectKind> = {
  'text.outline': 'text.outline',
  'text.emphasisMark': 'text.emphasisMark',
  'text.effect': 'text.effect',
  'text.modernEffects': 'text.modernEffects',
  'text.shadow': 'text.shadow',
  'text.leaderGlyphs': 'text.leaderGlyphs',
  'shape.effects': 'shape.effects',
  'image.filter': 'image.filter',
  'image.effects': 'image.effects',
};

/** The region codes `canvasSkippedEffects` reads in its `(region, area, index)` triples. */
const REGION = { body: 0, header: 1, footer: 2, noteSeparator: 3, note: 4 } as const;

/** The session method that walks the resident display page the canvas drew. */
interface CanvasEffectsSession {
  canvasSkippedEffects?(pageIndex: number, unavailable: Uint32Array): number;
}

/** The reference a picture primitive is drawn from; undefined for any other primitive. */
function pictureSource(primitive: DisplayPrimitive): string | undefined {
  if (primitive.kind === 'image') return primitive.relId || (primitive.relationshipId ?? '');
  if (primitive.kind !== 'shape' || primitive.fillPaint?.kind !== 'picture' || !primitive.geometryPath.length) return undefined;
  return primitive.fillPaint.pictureSrc ?? primitive.fillPaint.pictureRelId ?? '';
}

/** The `(region, area, index)` triples of the pictures on `page` drawn from `sources`. */
function picturePositions(page: DisplayPage, sources: ReadonlySet<string>): Uint32Array {
  const packed: number[] = [];
  const add = (region: number, area: number, primitives: readonly DisplayPrimitive[] | undefined) => {
    (primitives ?? []).forEach((primitive, index) => {
      const source = pictureSource(primitive);
      if (source !== undefined && sources.has(source)) packed.push(region, area, index);
    });
  };
  add(REGION.body, 0, page.primitives);
  add(REGION.header, 0, page.header?.primitives);
  add(REGION.footer, 0, page.footer?.primitives);
  (page.noteAreas ?? []).forEach((area, index) => {
    add(REGION.noteSeparator, index, area.separatorPrimitives);
    add(REGION.note, index, area.primitives);
  });
  return new Uint32Array(packed);
}

/** The kinds of effect the canvas left off the page `engine` drew. */
function canvasEffectKinds(engine: DocxPngEngine, unavailable: ReadonlySet<string>): PngSkippedEffectKind[] {
  const session = engine.session as CanvasEffectsSession;
  if (!session.canvasSkippedEffects) throw new Error('The session cannot report the effects the canvas leaves out.');
  const bits = session.canvasSkippedEffects(engine.page.pageIndex, picturePositions(engine.page, unavailable));
  return skippedEffectKinds().filter((_, bit) => bits & (1 << bit)).map((name) => {
    if (!Object.hasOwn(CANVAS_EFFECT_KINDS, name)) throw new Error(`The canvas PNG engine reported an unknown skipped effect: ${name}`);
    return CANVAS_EFFECT_KINDS[name as PngSkippedEffectKind];
  });
}

/** Whether display bytes are the stand-in the parser shows for a metafile it could not convert. */
function metafilePlaceholder(source: { bytes: Uint8Array; mimeType: string } | null): boolean {
  return source?.mimeType === 'image/svg+xml' && source.bytes.byteLength < 4096 && METAFILE_PLACEHOLDER.test(new TextDecoder().decode(source.bytes));
}

/** Paints one page with the editor's canvas backend and encodes it. */
export async function renderPng(engine: DocxPngEngine, options: Pick<DocxPngOptions, 'scale' | 'signal' | 'onDiagnostic' | 'strictPictures'>): Promise<Uint8Array> {
  const { width, height, scale } = pagePngRequest(engine.page, options);
  if (!canvasPngAvailable()) throw canvasUnavailable();
  options.signal?.throwIfAborted();
  const canvas = createCanvas(width, height);
  if (!canvas.getContext('2d')) throw canvasUnavailable();
  let refused: Error | undefined;
  const reasons = new Map<string, PngPictureUnavailableReason>();
  const pictures: PngUnavailablePicture[] = [];
  const unavailable = new Set<string>();
  const resolve = createCanvasImageResolver({ media: engine.media, onUnavailable: (id, reason) => reasons.set(id, reason) });
  await rasterizeDisplayPageToBackBuffer(canvas, { ...engine.page, background: engine.page.background ?? '#ffffff' }, {
    glyphCache: new GlyphCache({ provider: engine.outlineGlyph }),
    onGlyphFallback: () => { refused ??= fontUnavailable('A glyph on this page has no outline in the face layout measured it with.'); },
    resolveImage: (id) => {
      const source = engine.media(id);
      if (metafilePlaceholder(source)) {
        reasons.set(id, 'metafile');
        return null;
      }
      if (typeof Image === 'undefined' && (SVG_DATA_URL.test(id) || source?.mimeType === 'image/svg+xml')) {
        reasons.set(id, 'unsupported');
        return null;
      }
      return resolve(id);
    },
    onPictureUnavailable: (id) => {
      unavailable.add(id);
      pictures.push({ id: pictureId(id), reason: reasons.get(id) ?? 'missing' });
    },
  }, scale);
  if (refused) throw refused;
  if (pictures.length && (options.strictPictures ?? DEFAULT_STRICT_PICTURES)) throw strictPicturesError(pictures);
  const effects = canvasEffectKinds(engine, unavailable);
  options.signal?.throwIfAborted();
  const png = new Uint8Array(await (await encode(canvas)).arrayBuffer());
  if (pictures.length) options.onDiagnostic?.({ code: 'picture-unavailable', pictures });
  if (effects.length) options.onDiagnostic?.({ code: 'effects-skipped', page: engine.page.pageIndex, effects });
  return png;
}

/** A picture's reference as diagnostics name it: a `data:` URL by its header, never its payload. */
function pictureId(id: string): string {
  const comma = id.indexOf(',');
  return id.startsWith('data:') && comma >= 0 ? id.slice(0, comma) : id;
}

function createCanvas(width: number, height: number): PageCanvasBuffer {
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height);
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function encode(canvas: PageCanvasBuffer): Promise<Blob> {
  if ('convertToBlob' in canvas) return canvas.convertToBlob({ type: 'image/png' });
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('The canvas produced no PNG.'));
    }, 'image/png');
  });
}

export default renderPng;

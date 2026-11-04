import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { installPngFaces, type PngFontFace, type PngRunFont } from '../../../../shared/office-session/pngFonts';
import type { SlideDisplayList, SlidePrimitive } from '../types';
import { type CanvasImageResolver, paintSlide, slideBackingStore } from './canvas';

/** The font id layout gives text it measured with the browser's own fonts, which an export paints the same way. */
export const BROWSER_FONT_ID = 0xffffffff;

/** The engine's largest PNG side and area, in pixels. */
export const MAX_SLIDE_PNG_SIDE = 16_384;
export const MAX_SLIDE_PNG_PIXELS = 16_777_216;

/** `'slide'` paints white under the slide, as the editor shows it; `'transparent'` nothing; or a `#rrggbb` color. */
export type PptxPngBackground = 'slide' | 'transparent' | `#${string}`;

export interface SlideListToPngOptions {
  scale?: number;
  background?: PptxPngBackground;
  maxShadowPixels?: number;
  signal?: AbortSignal;
  fonts?: ReadonlyArray<SlideFontFace>;
  resolveImage?: CanvasImageResolver;
}

type ExportCanvas = HTMLCanvasElement | OffscreenCanvas;

/** Whether this runtime can paint and encode a PNG on a canvas: one with a 2d context, not only a document. */
export function canvasPngAvailable(): boolean {
  if (typeof OffscreenCanvas === 'function') return true;
  if (typeof document === 'undefined') return false;
  try { return document.createElement('canvas').getContext('2d') !== null; } catch { return false; }
}

export function canvasUnavailable(): PngExportError {
  return new PngExportError('canvas-unavailable', 'PNG export needs a canvas, and this environment has none. Use @betteroffice/pptx-native.');
}

/** Checks a PNG background the way the native renderer does. */
export function pngBackground(value: unknown): PptxPngBackground {
  if (value === undefined) return 'slide';
  if (value === 'slide' || value === 'transparent') return value;
  if (typeof value === 'string' && /^#[0-9a-fA-F]{6}$/.test(value)) return value as PptxPngBackground;
  throw new TypeError(`background must be "slide", "transparent", or a #rrggbb color, not ${JSON.stringify(value)}`);
}

/** Checks a slide export's options and size the way every engine does, and returns the PNG size. */
export function slidePngRequest(
  list: Pick<SlideDisplayList, 'width' | 'height'>, options: Pick<SlideListToPngOptions, 'scale' | 'background' | 'maxShadowPixels'>,
): { width: number; height: number; scale: number; background: PptxPngBackground } {
  const scale = options.scale ?? 1;
  if (typeof scale !== 'number' || !Number.isFinite(scale) || scale <= 0) throw new RangeError('scale must be finite and positive');
  const background = pngBackground(options.background);
  if (options.maxShadowPixels !== undefined && (!Number.isSafeInteger(options.maxShadowPixels) || options.maxShadowPixels < 0)) {
    throw new RangeError('invalid shadow pixel budget');
  }
  const { width, height } = slideBackingStore(list, 1, scale);
  if (width > MAX_SLIDE_PNG_SIDE || height > MAX_SLIDE_PNG_SIDE || width * height > MAX_SLIDE_PNG_PIXELS) {
    throw new PngExportError('canvas-too-large', `A ${width}x${height}px slide is over the ${MAX_SLIDE_PNG_SIDE}px side or ${MAX_SLIDE_PNG_PIXELS} pixel limit; lower the scale.`);
  }
  return { width, height, scale, background };
}

/**
 * Paints one laid-out slide with the editor's painter and encodes it. Faces in
 * `fonts` are installed for this export under private names; one that cannot
 * load rejects with `PngExportError` rather than paint a substitute.
 */
export async function slideListToPng(list: SlideDisplayList, options: SlideListToPngOptions = {}): Promise<Uint8Array> {
  const { width, height, scale, background } = slidePngRequest(list, options);
  options.signal?.throwIfAborted();
  const fonts = slideFonts(list);
  const missing = fonts.find((font) => font.id !== undefined && font.id !== BROWSER_FONT_ID && !slideFace(options.fonts ?? [], font));
  if (missing) {
    throw new PngExportError('font-unavailable', `The slide's text was laid out with a ${JSON.stringify(missing.family)} face this export does not have.`);
  }
  const faces = await installPngFaces({
    prefix: 'bo-pptx-png', fonts, faces: options.fonts ?? [], choose: slideFace, subject: 'slide',
    canvasFont: (font, family) => `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}16px ${JSON.stringify(family)}`,
  });
  try {
    options.signal?.throwIfAborted();
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
    if (!ctx) throw new PngExportError('canvas-unavailable', 'PNG export needs a 2d canvas context.');
    await paintSlide(ctx, list, 1, scale, {
      resolveImage: options.resolveImage, maxShadowPixels: options.maxShadowPixels, fontFamily: faces.family,
    });
    if (background !== 'transparent') {
      ctx.save();
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.globalCompositeOperation = 'destination-over';
      ctx.fillStyle = background === 'slide' ? '#ffffff' : background;
      ctx.fillRect(0, 0, width, height);
      ctx.restore();
    }
    options.signal?.throwIfAborted();
    return new Uint8Array(await (await encode(canvas)).arrayBuffer());
  } finally { faces.release(); }
}

function collect(primitives: readonly SlidePrimitive[], fonts: Map<string, PngRunFont>): void {
  for (const primitive of primitives) {
    if (primitive.kind === 'chart' || primitive.kind === 'table') collect(primitive.primitives, fonts);
    if (primitive.kind !== 'textBox') continue;
    for (const line of primitive.lines) {
      for (const run of line.runs) {
        const font = { family: run.fontFamily, bold: run.bold === true, italic: run.italic === true, id: run.fontId };
        fonts.set(JSON.stringify(font), font);
      }
    }
  }
}

/** The families and styles a slide's text paints with. */
export function slideFonts(list: SlideDisplayList): PngRunFont[] {
  const fonts = new Map<string, PngRunFont>();
  collect(list.primitives, fonts);
  return [...fonts.values()];
}

const familyName = (family: string) => family.trim().toLowerCase();

/** A face for an export, with the engine's id when it is one the presentation registered. */
export type SlideFontFace = PngFontFace & { id?: number };

/** The face layout measured a run with, by the engine's id; else a supplied face by family, dropping bold before italic. */
export function slideFace(faces: ReadonlyArray<SlideFontFace>, font: PngRunFont): SlideFontFace | undefined {
  const measured = font.id === undefined ? undefined : faces.find((face) => face.id === font.id);
  if (measured) return measured;
  for (const [bold, italic] of [[font.bold, font.italic], [font.bold, false], [false, font.italic], [false, false]]) {
    let found: SlideFontFace | undefined;
    for (const face of faces) {
      if (face.id === undefined && familyName(face.family) === familyName(font.family) && (face.bold ?? false) === bold && (face.italic ?? false) === italic) found = face;
    }
    if (found) return found;
  }
  return undefined;
}

function createCanvas(width: number, height: number): ExportCanvas {
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height);
  if (typeof document === 'undefined') throw canvasUnavailable();
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function encode(canvas: ExportCanvas): Promise<Blob> {
  if ('convertToBlob' in canvas) return canvas.convertToBlob({ type: 'image/png' });
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new PngExportError('canvas-too-large', 'The canvas produced no PNG.'));
    }, 'image/png');
  });
}

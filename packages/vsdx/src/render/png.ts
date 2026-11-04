import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { installPngFaces, pngFontKey } from '../../../../shared/office-session/pngFonts';
import type { PageDisplayList, PagePrimitive, TextRun, VsdxFontFace } from '../types';
import { canvasFontFor, effectiveDprForSurface, paintPage, type TextFace } from './canvas';

export interface PageToPngOptions {
  scale?: number;
  signal?: AbortSignal;
  fonts?: ReadonlyArray<VsdxFontFace>;
  media?: (assetId: string) => Uint8Array;
}

type ExportCanvas = HTMLCanvasElement | OffscreenCanvas;
interface RunFont { family: string; bold: boolean; italic: boolean; }

export async function pageToPng(list: PageDisplayList, options: PageToPngOptions = {}): Promise<Uint8Array> {
  const scale = options.scale ?? 1;
  const width = Math.max(1, Math.ceil(list.width * scale));
  const height = Math.max(1, Math.ceil(list.height * scale));
  if (effectiveDprForSurface(width, height, 1) < 1) {
    throw new PngExportError('canvas-too-large', `A ${width}x${height}px page is over the canvas limit; lower the scale.`);
  }
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) throw new PngExportError('canvas-unavailable', 'PNG export needs a 2d canvas context.');
  const installed = await installFaces(list, options.fonts ?? []);
  const bitmaps: ImageBitmap[] = [];
  try {
    await paintPage(ctx, list, 1, scale, { signal: options.signal, outlineMissingImages: true, textFace: installed.face, resolveImage: options.media && (assetId => decodeImage(assetId, options.media!, bitmaps)) });
    options.signal?.throwIfAborted();
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    ctx.restore();
    return new Uint8Array(await (await encode(canvas)).arrayBuffer());
  } finally {
    for (const bitmap of bitmaps) bitmap.close();
    installed.release();
  }
}

function createCanvas(width: number, height: number): ExportCanvas {
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height);
  if (typeof document === 'undefined') throw new PngExportError('canvas-unavailable', 'PNG export needs a canvas, and this environment has none. Use @betteroffice/vsdx-native.');
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function encode(canvas: ExportCanvas): Promise<Blob> {
  if ('convertToBlob' in canvas) return canvas.convertToBlob({ type: 'image/png' });
  return new Promise((resolve, reject) => {
    canvas.toBlob(blob => { if (blob) resolve(blob); else reject(new PngExportError('canvas-too-large', 'The canvas produced no PNG.')); }, 'image/png');
  });
}

async function decodeImage(assetId: string, media: (assetId: string) => Uint8Array, bitmaps: ImageBitmap[]): Promise<ImageBitmap | null> {
  try {
    const bitmap = await createImageBitmap(new Blob([media(assetId).slice()]));
    bitmaps.push(bitmap);
    return bitmap;
  } catch { return null; }
}

function collectRunFonts(primitives: PagePrimitive[], fonts: Map<string, RunFont>): void {
  for (const primitive of primitives) {
    if (primitive.kind === 'group') collectRunFonts(primitive.primitives, fonts);
    if (primitive.kind !== 'textBox') continue;
    for (const paragraph of primitive.paragraphs) {
      for (const { family, bold, italic } of paragraph.runs) fonts.set(pngFontKey({ family, bold, italic }), { family, bold, italic });
    }
  }
}

/** The registered face layout measured a run with: the family, then Arial, then sans-serif, each dropping styles in layout's order. */
export function layoutFace(faces: ReadonlyArray<VsdxFontFace>, font: RunFont): VsdxFontFace | undefined {
  for (const family of [font.family, 'Arial', 'sans-serif']) {
    for (const [bold, italic] of [[font.bold, font.italic], [font.bold, false], [false, font.italic], [false, false]]) {
      let found: VsdxFontFace | undefined;
      for (const face of faces) if (face.family === family && (face.bold ?? false) === bold && (face.italic ?? false) === italic) found = face;
      if (found) return found;
    }
  }
  return undefined;
}

/** Underline position and thickness in ems from the face's post table, or the raster defaults. */
export function underlineMetrics(bytes: Uint8Array): { position: number; thickness: number } {
  const fallback = { position: -0.1, thickness: 0.05 };
  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const base = view.getUint32(0) === 0x74746366 ? view.getUint32(12) : 0;
    const count = view.getUint16(base + 4);
    let head = -1, post = -1;
    for (let index = 0; index < count; index++) {
      const record = base + 12 + index * 16;
      const tag = view.getUint32(record);
      if (tag === 0x68656164) head = view.getUint32(record + 8);
      if (tag === 0x706f7374) post = view.getUint32(record + 8);
    }
    if (head < 0 || post < 0) return fallback;
    const em = view.getUint16(head + 18);
    if (!em) return fallback;
    return { position: view.getInt16(post + 8) / em, thickness: view.getInt16(post + 10) / em };
  } catch { return fallback; }
}

interface InstalledFaces { face(run: TextRun): TextFace | undefined; release(): void; }

async function installFaces(list: PageDisplayList, faces: ReadonlyArray<VsdxFontFace>): Promise<InstalledFaces> {
  const fonts = new Map<string, RunFont>();
  collectRunFonts(list.primitives, fonts);
  const installed = await installPngFaces({
    prefix: 'bo-vsdx-png', fonts: fonts.values(), faces, choose: layoutFace, subject: 'page',
    canvasFont: (font, family) => canvasFontFor({ ...font, family }, 16),
  });
  const metrics = new Map<VsdxFontFace, TextFace>();
  return {
    face: (run) => {
      const family = installed.family(run);
      const face = installed.face(run);
      if (!family || !face) return undefined;
      let text = metrics.get(face);
      if (!text) {
        const underline = underlineMetrics(face.bytes);
        text = { family, underlinePosition: underline.position, underlineThickness: underline.thickness };
        metrics.set(face, text);
      }
      return text;
    },
    release: installed.release,
  };
}

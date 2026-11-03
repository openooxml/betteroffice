import { PngExportError } from '../../../../shared/office-session/pngExportError';
import type { PageDisplayList, PagePrimitive, VsdxFontFace } from '../types';
import { canvasFontFor, effectiveDprForSurface, paintPage } from './canvas';

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
    await paintPage(ctx, list, 1, scale, { signal: options.signal, outlineMissingImages: true, fontFamily: installed.alias, resolveImage: options.media && (assetId => decodeImage(assetId, options.media!, bitmaps)) });
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

function fontSet(): FontFaceSet | undefined {
  return typeof document !== 'undefined' ? document.fonts : (globalThis as { fonts?: FontFaceSet }).fonts;
}

function collectRunFonts(primitives: PagePrimitive[], fonts: Map<string, RunFont>): void {
  for (const primitive of primitives) {
    if (primitive.kind === 'group') collectRunFonts(primitive.primitives, fonts);
    if (primitive.kind !== 'textBox') continue;
    for (const paragraph of primitive.paragraphs) {
      for (const { family, bold, italic } of paragraph.runs) fonts.set(JSON.stringify([family, bold, italic]), { family, bold, italic });
    }
  }
}

const familyName = (family: string) => family.trim().replace(/^(['"])(.*)\1$/, '$2').toLowerCase();

function unavailable(family: string): PngExportError {
  return new PngExportError('font-unavailable', `The font "${family}" could not load, so the page was not exported with a substitute.`);
}

let exports = 0;

interface InstalledFaces { alias(family: string): string; release(): void; }

async function installFaces(list: PageDisplayList, faces: ReadonlyArray<VsdxFontFace>): Promise<InstalledFaces> {
  const fonts = new Map<string, RunFont>();
  collectRunFonts(list.primitives, fonts);
  const used = new Set([...fonts.values()].map(font => familyName(font.family)));
  const set = fontSet();
  const aliases = new Map<string, string>();
  const added: FontFace[] = [];
  const release = () => { for (const face of added) set?.delete(face); };
  const id = ++exports;
  try {
    for (const face of faces) {
      const name = familyName(face.family);
      if (!used.has(name)) continue;
      if (!set || typeof FontFace === 'undefined') throw unavailable(face.family);
      if (!aliases.has(name)) aliases.set(name, `bo-vsdx-png-${id}-${aliases.size}`);
      const loaded = await new FontFace(aliases.get(name)!, face.bytes.slice().buffer as ArrayBuffer, {
        weight: face.bold ? '700' : '400', style: face.italic ? 'italic' : 'normal',
      }).load().catch(() => { throw unavailable(face.family); });
      set.add(loaded);
      added.push(loaded);
    }
    const alias = (family: string) => aliases.get(familyName(family)) ?? family;
    for (const font of fonts.values()) {
      const request = canvasFontFor({ ...font, family: alias(font.family) }, 16);
      const matched = await set?.load(request).catch(() => { throw unavailable(font.family); });
      if (aliases.has(familyName(font.family)) && !matched?.length) throw unavailable(font.family);
    }
    return { alias, release };
  } catch (error) {
    release();
    throw error;
  }
}

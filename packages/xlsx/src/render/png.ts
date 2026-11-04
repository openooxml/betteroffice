import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { fontUnavailable, installPngFaces, pngFontSet, type PngFontFace, type PngRunFont } from '../../../../shared/office-session/pngFonts';
import type { DisplayList, TextCmd } from '../display-list/types';
import { MAX_PNG_PIXELS, MAX_PNG_SIDE } from '../wasm/pngResolve';
import { paintDisplayList } from './canvas2d';
import { fontString } from './fontString';

export interface DisplayListToPngOptions {
  /** Device pixels per display-list pixel. */
  dpr: number;
  /** PNG size in pixels, as the engine sized it. */
  width: number;
  height: number;
  signal?: AbortSignal;
  /** Faces to paint with, matched to the sheet's fonts by family and style. */
  fonts?: ReadonlyArray<PngFontFace>;
}

type ExportCanvas = HTMLCanvasElement | OffscreenCanvas;

/** Whether this runtime can paint and encode a PNG on a canvas: one with a 2d context, not only a document. */
export function canvasPngAvailable(): boolean {
  if (typeof OffscreenCanvas === 'function') return true;
  if (typeof document === 'undefined') return false;
  try { return document.createElement('canvas').getContext('2d') !== null; } catch { return false; }
}

/** Paints a sheet display list with the editor's painter on a white canvas and encodes it. */
export async function displayListToPng(list: DisplayList, options: DisplayListToPngOptions): Promise<Uint8Array> {
  const { width, height, dpr } = options;
  if (width > MAX_PNG_SIDE || height > MAX_PNG_SIDE || width * height > MAX_PNG_PIXELS) {
    throw new PngExportError('canvas-too-large', `A ${width}x${height}px PNG is over the ${MAX_PNG_SIDE}px side or ${MAX_PNG_PIXELS} pixel limit; lower the scale or set maxWidth and maxHeight.`);
  }
  options.signal?.throwIfAborted();
  const faces = await installPngFaces({
    prefix: 'bo-xlsx-png', fonts: sheetFonts(list), faces: options.fonts ?? [], choose: sheetFace, subject: 'sheet',
    canvasFont: (font, family) => fontString({ op: 'text', x: 0, y: 0, text: '', color: '#000000', fontSize: 12, fontFamily: family, bold: font.bold, italic: font.italic }),
  });
  try {
    await loadChartFonts(list);
    options.signal?.throwIfAborted();
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
    if (!ctx) throw new PngExportError('canvas-unavailable', 'PNG export needs a 2d canvas context.');
    paintDisplayList(ctx, withFaces(list, faces.family), dpr);
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    ctx.restore();
    options.signal?.throwIfAborted();
    return new Uint8Array(await (await encode(canvas)).arrayBuffer());
  } finally { faces.release(); }
}

/** The error for a runtime that can draw the PNG neither on a canvas nor with a rasterizer. */
export function canvasUnavailable(detail = 'PNG export needs a canvas, and this environment has none.'): PngExportError {
  return new PngExportError('canvas-unavailable', `${detail} Use @betteroffice/xlsx-native.`);
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

const runFont = (cmd: TextCmd): PngRunFont => ({ family: cmd.fontFamily!, bold: cmd.bold === true, italic: cmd.italic === true });
const isCellText = (cmd: DisplayList['commands'][number]): cmd is TextCmd => cmd.op === 'text' && !!cmd.fontFamily && !cmd.chart;

/** The families and styles the sheet's cell text paints with. */
export function sheetFonts(list: DisplayList): PngRunFont[] {
  const fonts = new Map<string, PngRunFont>();
  for (const cmd of list.commands) if (isCellText(cmd)) fonts.set(JSON.stringify(runFont(cmd)), runFont(cmd));
  return [...fonts.values()];
}

const familyName = (family: string) => family.trim().toLowerCase();

/** The supplied face for a cell font: its family, dropping bold before italic, as the editor resolves faces. */
export function sheetFace(faces: ReadonlyArray<PngFontFace>, font: PngRunFont): PngFontFace | undefined {
  for (const [bold, italic] of [[font.bold, font.italic], [font.bold, false], [false, font.italic], [false, false]]) {
    let found: PngFontFace | undefined;
    for (const face of faces) {
      if (familyName(face.family) === familyName(font.family) && (face.bold ?? false) === bold && (face.italic ?? false) === italic) found = face;
    }
    if (found) return found;
  }
  return undefined;
}

function withFaces(list: DisplayList, family: (font: PngRunFont) => string | undefined): DisplayList {
  let changed = false;
  const commands = list.commands.map((cmd) => {
    if (!isCellText(cmd)) return cmd;
    const alias = family(runFont(cmd));
    if (!alias) return cmd;
    changed = true;
    return { ...cmd, fontFamily: alias };
  });
  return changed ? { ...list, commands } : list;
}

/** Chart text names CSS font lists of its own; wait for them as the editor would. */
async function loadChartFonts(list: DisplayList): Promise<void> {
  const set = pngFontSet();
  if (!set || typeof set.load !== 'function') return;
  const fonts = new Map<string, string>();
  for (const cmd of list.commands) if (cmd.op === 'text' && (cmd.chart || !cmd.fontFamily)) fonts.set(fontString(cmd), cmd.fontFamily ?? 'sans-serif');
  for (const [font, family] of fonts) await set.load(font).catch(() => { throw fontUnavailable(family, 'sheet'); });
}

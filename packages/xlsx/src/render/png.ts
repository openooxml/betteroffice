import { PngExportError } from '../../../../shared/office-session/pngExportError';
import { fontUnavailable, installPngFaces, pngFontKey, pngFontSet, type PngFontFace, type PngRunFont } from '../../../../shared/office-session/pngFonts';
import { textScripts } from '../../../../shared/office-session/textScripts';
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
  /** Faces that paint the characters a font's face lacks, in the order the native engine tries them. */
  fallbacks?: (font: PngRunFont) => ReadonlyArray<PngFontFace>;
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
    prefix: 'bo-xlsx-png', fonts: sheetFonts(list), faces: options.fonts ?? [], choose: sheetFace, fallbacks: options.fallbacks, subject: 'sheet',
    canvasFont: (font, family) => fontString({ op: 'text', x: 0, y: 0, text: '', color: '#000000', fontSize: 12, fontFamily: family, bold: font.bold, italic: font.italic }),
  });
  try {
    await loadChartFonts(list);
    options.signal?.throwIfAborted();
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
    if (!ctx) throw canvasUnavailable('PNG export needs a 2d canvas context.');
    paintDisplayList(ctx, withFaces(list, faces.families), dpr);
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

/** The error for a runtime without a canvas. */
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

const GENERIC = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-sans-serif', 'ui-serif', 'ui-monospace']);

/** The family a text command asks for first: the cell's family, or the head of a chart's CSS font list unless it is generic. */
function primaryFamily(cmd: TextCmd): string | undefined {
  if (!cmd.fontFamily) return undefined;
  if (!cmd.chart) return cmd.fontFamily;
  const head = cmd.fontFamily.split(',')[0]?.trim().replace(/^(['"])(.*)\1$/, '$2');
  return head && !GENERIC.has(head.toLowerCase()) ? head : undefined;
}

function runFont(cmd: TextCmd): PngRunFont | undefined {
  const family = primaryFamily(cmd);
  return family === undefined ? undefined : { family, bold: cmd.bold === true, italic: cmd.italic === true };
}

/** A font the sheet's text asks for, with the script buckets of one text in it. */
export interface SheetFontRequest {
  font: PngRunFont;
  scripts: string[];
}

/**
 * Each font and the scripts of the text set in it, in display-list order and once each, as the native engine walks
 * the list to resolve the faces it draws with.
 */
export function sheetFontRequests(list: DisplayList): SheetFontRequest[] {
  const requests = new Map<string, SheetFontRequest>();
  for (const cmd of list.commands) {
    const font = cmd.op === 'text' ? runFont(cmd) : undefined;
    if (!font || cmd.op !== 'text') continue;
    const scripts = textScripts(cmd.text);
    const key = JSON.stringify([pngFontKey(font), scripts]);
    if (!requests.has(key)) requests.set(key, { font, scripts });
  }
  return [...requests.values()];
}

/** The families and styles the sheet's cell and chart text paints with. */
export function sheetFonts(list: DisplayList): PngRunFont[] {
  const fonts = new Map<string, PngRunFont>();
  for (const cmd of list.commands) {
    const font = cmd.op === 'text' ? runFont(cmd) : undefined;
    if (font) fonts.set(JSON.stringify(font), font);
  }
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

function withFaces(list: DisplayList, families: (font: PngRunFont) => string[]): DisplayList {
  let changed = false;
  const commands = list.commands.map((cmd) => {
    const font = cmd.op === 'text' ? runFont(cmd) : undefined;
    const aliases = font ? families(font) : [];
    if (!aliases.length || cmd.op !== 'text') return cmd;
    changed = true;
    const head = aliases.map((alias) => `"${alias}"`).join(', ');
    return {
      ...cmd, fontFamily: cmd.chart ? `"${aliases[0]}", ${cmd.fontFamily}` : aliases[0],
      cssFontFamily: `${head}, ${cmd.chart ? cmd.fontFamily : 'sans-serif'}`,
    };
  });
  return changed ? { ...list, commands } : list;
}

/** Text with no family of its own to supply a face for (generic chart lists, unnamed fonts); wait for it as the editor would. */
async function loadChartFonts(list: DisplayList): Promise<void> {
  const set = pngFontSet();
  if (!set || typeof set.load !== 'function') return;
  const fonts = new Map<string, string>();
  for (const cmd of list.commands) if (cmd.op === 'text' && !runFont(cmd)) fonts.set(fontString(cmd), cmd.fontFamily ?? 'sans-serif');
  for (const [font, family] of fonts) await set.load(font).catch(() => { throw fontUnavailable(family, 'sheet'); });
}

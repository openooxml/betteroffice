import type { DisplayPage, DisplayPrimitive, GlyphRunPrimitive } from '@betteroffice/docx/layout/render';
import type { DocxEchoRect } from '@betteroffice/docx/worker';

type Rgb = readonly [number, number, number];

const WHITE: Rgb = [255, 255, 255];
const EDGE = 0.5;

function parseRgba(value: string | undefined): readonly [number, number, number, number] | null {
  if (!value) return null;
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(value.trim());
  if (hex) {
    const digits = hex[1]!.length === 3 ? [...hex[1]!].map((digit) => digit + digit).join('') : hex[1]!;
    const [r, g, b] = [0, 2, 4].map((offset) => parseInt(digits.slice(offset, offset + 2), 16));
    return [r!, g!, b!, digits.length === 8 ? parseInt(digits.slice(6), 16) / 255 : 1];
  }
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/i.exec(value.trim());
  if (!rgb) return null;
  const alpha = rgb[4] === undefined ? 1 : rgb[4].endsWith('%') ? parseFloat(rgb[4]) / 100 : parseFloat(rgb[4]);
  return [Number(rgb[1]), Number(rgb[2]), Number(rgb[3]), Math.min(1, Math.max(0, alpha))];
}

function parseColor(value: string | undefined): Rgb | null {
  const rgba = parseRgba(value);
  return rgba && rgba[3] >= 0.5 ? [rgba[0], rgba[1], rgba[2]] : null;
}

/** A colour painted over `under` at `opacity`, alpha composited. */
function over(value: string | undefined, under: Rgb, opacity = 1): Rgb | null {
  const rgba = parseRgba(value);
  if (!rgba) return null;
  const alpha = rgba[3] * opacity;
  return [0, 1, 2].map((channel) => Math.round(rgba[channel]! * alpha + under[channel]! * (1 - alpha))) as unknown as Rgb;
}

function luminance([r, g, b]: Rgb): number {
  const channel = (value: number) => {
    const c = value / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

function contrast(a: Rgb, b: Rgb): number {
  const [light, dark] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (light! + 0.05) / (dark! + 0.05);
}

function covers(box: { x: number; y: number; w: number; h: number }, x: number, y: number): boolean {
  return x >= box.x && x < box.x + box.w && y >= box.y && y < box.y + box.h;
}

/** Whether a possibly rotated box may paint at the point: its bounds, or the circle any rotation stays in. */
function mayCover(box: { x: number; y: number; w: number; h: number }, rotation: number | undefined, x: number, y: number): boolean {
  if (!rotation || rotation % 360 === 0) return covers(box, x, y);
  const radius = Math.hypot(box.w, box.h) / 2;
  return Math.hypot(x - (box.x + box.w / 2), y - (box.y + box.h / 2)) <= radius;
}

function clipped(primitive: DisplayPrimitive, x: number, y: number): boolean {
  const group = primitive.clipGroup;
  if (group?.opacity !== undefined && group.opacity < 0.5) return true;
  const clip = group?.clip;
  if (!clip) return false;
  if (clip.x !== undefined && (x < clip.x || (clip.w !== undefined && x >= clip.x + clip.w))) return true;
  return clip.y !== undefined && (y < clip.y || (clip.h !== undefined && y >= clip.y + clip.h));
}

function* paintOrder(page: DisplayPage): Generator<DisplayPrimitive> {
  const watermarks = Math.min(page.watermarkPrimitiveCount ?? 0, page.primitives.length);
  yield* page.primitives.slice(0, watermarks);
  if (page.header) yield* page.header.primitives;
  if (page.footer) yield* page.footer.primitives;
  yield* page.primitives.slice(watermarks);
  for (const area of page.noteAreas ?? []) yield* area.primitives ?? [];
}

function runExtent(run: GlyphRunPrimitive): { left: number; right: number; baseline: number } | null {
  const first = run.glyphs[0];
  const last = run.glyphs.at(-1);
  if (!first || !last) return null;
  return { left: first.x, right: last.x + (last.advance ?? 0), baseline: first.y };
}

/**
 * Caret paint for a page-local caret rect: the colour of the text the caret
 * sits in, unless a known solid fill under the caret would hide it; then black
 * or white by that fill. Shapes and images make the fill unknown: the text
 * colour is used as is and `halo` gives the caret a contrasting outline, so it
 * shows on whatever is painted there.
 */
export function peerlessCaretPaint(page: DisplayPage | undefined, caret: DocxEchoRect): { color: string; halo?: string } {
  if (!page) return { color: '#000' };
  const x = caret.x;
  const y = caret.y + caret.height / 2;
  const paintX = caret.x + Math.max(1, caret.width) / 2;
  let fill: Rgb = over(page.background, WHITE) ?? WHITE;
  let known = true;
  let before: GlyphRunPrimitive | undefined;
  let after: GlyphRunPrimitive | undefined;
  const paint = (value: string | undefined, opacity = 1) => {
    const color = over(value, fill, opacity);
    if (color) {
      fill = color;
      if (opacity * (parseRgba(value)?.[3] ?? 0) >= 0.5) known = true;
    } else if (value && !/^(transparent|none)$/i.test(value.trim())) known = false;
  };
  for (const primitive of paintOrder(page)) {
    if (primitive.kind === 'rect') {
      if (covers(primitive, paintX, y) && !clipped(primitive, paintX, y)) paint(primitive.fill, primitive.opacity ?? 1);
    } else if (primitive.kind === 'decoration') {
      if (primitive.deco === 'highlight' && covers(primitive, paintX, y) && !clipped(primitive, paintX, y)) paint(primitive.color);
    } else if (primitive.kind === 'shape' || primitive.kind === 'image') {
      const rotation = primitive.kind === 'shape' ? primitive.transform?.rotation : primitive.rotationDeg;
      if (mayCover(primitive, rotation, paintX, y) && !clipped(primitive, paintX, y)) known = false;
    } else if (primitive.kind === 'glyphRun' && !primitive.listMarker && !primitive.hidden) {
      const extent = runExtent(primitive);
      if (!extent || extent.baseline < caret.y || extent.baseline > caret.y + caret.height) continue;
      if (x < extent.left - EDGE || x > extent.right + EDGE) continue;
      if (x > extent.left + EDGE) before = primitive;
      else after ??= primitive;
    }
  }
  const run = before ?? after;
  const text = parseColor(run?.color);
  const fallback = luminance(fill) < 0.18 ? '#fff' : '#000';
  if (!known) {
    const color = run && text ? run.color : fallback;
    return { color, halo: luminance(text && run ? text : parseColor(color)!) < 0.18 ? '#fff' : '#000' };
  }
  return { color: run && text && contrast(text, fill) >= 2 ? run.color : fallback };
}

/** The caret colour alone, for overlays that draw no halo. */
export function peerlessCaretColor(page: DisplayPage | undefined, caret: DocxEchoRect): string {
  return peerlessCaretPaint(page, caret).color;
}

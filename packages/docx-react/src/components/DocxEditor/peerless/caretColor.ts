import type { DisplayPage, DisplayPrimitive, GlyphRunPrimitive } from '@betteroffice/docx/layout/render';
import type { DocxEchoRect } from '@betteroffice/docx/worker';

type Rgb = readonly [number, number, number];

const WHITE: Rgb = [255, 255, 255];
const EDGE = 0.5;

function parseColor(value: string | undefined): Rgb | null {
  if (!value) return null;
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value.trim());
  if (hex) {
    const digits = hex[1]!.length === 3 ? [...hex[1]!].map((digit) => digit + digit).join('') : hex[1]!;
    return [0, 2, 4].map((offset) => parseInt(digits.slice(offset, offset + 2), 16)) as unknown as Rgb;
  }
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/i.exec(value.trim());
  if (!rgb) return null;
  if (rgb[4] !== undefined && parseFloat(rgb[4]) < (rgb[4].endsWith('%') ? 50 : 0.5)) return null;
  return [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])];
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
 * Caret paint colour for a page-local caret rect: the colour of the text the
 * caret sits in, as long as it reads against the fill under the caret; else
 * black or white by that fill (the page is white when nothing is painted).
 */
export function peerlessCaretColor(page: DisplayPage | undefined, caret: DocxEchoRect): string {
  if (!page) return '#000';
  const x = caret.x;
  const y = caret.y + caret.height / 2;
  const paintX = caret.x + Math.max(1, caret.width) / 2;
  let fill: Rgb = parseColor(page.background) ?? WHITE;
  let before: GlyphRunPrimitive | undefined;
  let after: GlyphRunPrimitive | undefined;
  for (const primitive of paintOrder(page)) {
    if (primitive.kind === 'rect') {
      if ((primitive.opacity ?? 1) >= 0.5 && covers(primitive, paintX, y)) fill = parseColor(primitive.fill) ?? fill;
    } else if (primitive.kind === 'decoration') {
      if (primitive.deco === 'highlight' && covers(primitive, paintX, y)) fill = parseColor(primitive.color) ?? fill;
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
  if (run && text && contrast(text, fill) >= 2) return run.color;
  return luminance(fill) < 0.18 ? '#fff' : '#000';
}

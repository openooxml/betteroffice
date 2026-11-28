import type { Paint, PositionedTextLine, SlideDisplayList, SlidePrimitive } from '@betteroffice/pptx';
import { transformedRect, type CommandRect } from '@betteroffice/pptx/worker';

type Rgb = readonly [number, number, number];

const WHITE = '#ffffff';

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

/** A colour painted over `under`, alpha composited. */
function over(value: string | undefined, under: Rgb): Rgb | null {
  const rgba = parseRgba(value);
  if (!rgba) return null;
  const alpha = rgba[3];
  return [0, 1, 2].map((channel) => Math.round(rgba[channel]! * alpha + under[channel]! * (1 - alpha))) as unknown as Rgb;
}

function paintColor(paint: Paint | undefined, under: Rgb): Rgb | null {
  if (!paint) return null;
  if (paint.kind === 'solid') return over(paint.color, under);
  const stops = paint.stops.map((stop) => over(stop.color, under)).filter((stop): stop is Rgb => stop !== null);
  if (!stops.length) return null;
  return [0, 1, 2].map((channel) => Math.round(stops.reduce((sum, stop) => sum + stop[channel]!, 0) / stops.length)) as unknown as Rgb;
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

function hex([r, g, b]: Rgb): string {
  return '#' + [r, g, b].map((channel) => channel.toString(16).padStart(2, '0')).join('');
}

/** `estimate` is the best guess at the painted colour, or `null` when unknown. */
interface Fill { color: Rgb; solid: boolean; estimate: Rgb | null }

function fillOf(paint: Paint | undefined, under: Fill): Fill | null {
  const color = paintColor(paint, under.color);
  if (!color || !paint) return null;
  const opaque = paint.kind === 'solid' ? (parseRgba(paint.color)?.[3] ?? 0) >= 1
    : paint.stops.every((stop) => (parseRgba(stop.color)?.[3] ?? 0) >= 1);
  const estimate = opaque ? color : under.estimate && paintColor(paint, under.estimate);
  return { color, solid: paint.kind === 'solid' && (opaque || under.solid), estimate };
}

function fillAt(primitives: readonly SlidePrimitive[], parents: readonly SlidePrimitive[], x: number, y: number, fill: Fill): Fill {
  for (const primitive of primitives) {
    const transforms = [...parents, primitive];
    if (primitive.kind === 'table' || primitive.kind === 'chart') {
      fill = fillAt(primitive.primitives, transforms, x, y, fill);
      continue;
    }
    if (primitive.kind !== 'shape' && primitive.kind !== 'image') continue;
    const paint = primitive.kind === 'shape' ? fillOf(primitive.fill, fill) : null;
    if (primitive.kind === 'shape' && !paint) continue;
    const box = transformedRect(transforms, { x: primitive.x, y: primitive.y, width: primitive.w, height: primitive.h });
    if (!(x >= box.x && x < box.x + box.width && y >= box.y && y < box.y + box.height)) continue;
    const upright = transforms.every(({ transform }) => !transform?.rotationDeg && !transform?.flipH && !transform?.flipV);
    const exact = primitive.kind === 'shape' && primitive.geometry === 'rect' && upright;
    fill = paint && exact ? paint : { color: fill.color, solid: false, estimate: upright ? paint?.estimate ?? null : null };
  }
  return fill;
}

const BLACK = '#000000';

function contrasting(color: Rgb): string {
  return luminance(color) < 0.18 ? WHITE : BLACK;
}

function matching(color: Rgb): string {
  return luminance(color) < 0.18 ? BLACK : WHITE;
}

/**
 * Colours at a slide-space caret rect. `color` is the text colour at the caret
 * (its run, or the insertion colour on an empty line) when it reads against a
 * known solid fill, else black or white by that fill. `caret` is black or white
 * against the best estimate of the fill behind it. When the fill is unknown,
 * it follows the text shade if known, or contrasts with the last known fill.
 */
export function pptxCaretColors(list: SlideDisplayList | null | undefined, line: PositionedTextLine | null | undefined,
  position: number | null | undefined, caret: CommandRect | null | undefined,
  insertion?: string | null): { color: string; caret: string; background: string } {
  const white: Fill = { color: [255, 255, 255], solid: true, estimate: [255, 255, 255] };
  let under: Fill = fillOf(list?.background, white) ?? white;
  if (list && caret) under = fillAt(list.primitives, [], caret.x + caret.width / 2, caret.y + caret.height / 2, under);
  const fill = under.color;
  const runs = line?.runs ?? [];
  const run = position === null || position === undefined ? undefined
    : runs.find((candidate) => candidate.start < position && position <= candidate.end)
      ?? runs.find((candidate) => candidate.start <= position && position <= candidate.end);
  const textColor = run?.color ?? insertion ?? undefined;
  const text = parseColor(textColor);
  const fallback = contrasting(fill);
  const bar = under.estimate ? ink(luminance(under.estimate)) : text ? matching(text) : fallback;
  if (!under.solid) {
    const color = text ? textColor! : fallback;
    const opposite = (text ? matching(text) : fallback) === WHITE ? BLACK : WHITE;
    return { color, caret: bar, background: opposite };
  }
  const background = hex(fill);
  if (!text || contrast(text, fill) < 2) return { color: fallback, caret: bar, background };
  return { color: textColor!, caret: bar, background };
}

type PixelRead = (source: HTMLCanvasElement, left: number, top: number, width: number, height: number) => Uint8ClampedArray | undefined;

const scratches = new WeakMap<Document, { canvas: HTMLCanvasElement; context: CanvasRenderingContext2D } | null>();

/** Copies the region into a small read-optimised canvas, so the slide canvas itself is never read back. */
const readThroughScratch: PixelRead = (source, left, top, width, height) => {
  const document = source.ownerDocument;
  try {
    let scratch = scratches.get(document);
    if (scratch === undefined) {
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d', { willReadFrequently: true });
      scratch = context ? { canvas, context } : null;
      scratches.set(document, scratch);
    }
    if (!scratch) return undefined;
    if (scratch.canvas.width < width) scratch.canvas.width = width;
    if (scratch.canvas.height < height) scratch.canvas.height = height;
    scratch.context.clearRect(0, 0, width, height);
    scratch.context.drawImage(source, left, top, width, height, 0, 0, width, height);
    return scratch.context.getImageData(0, 0, width, height).data;
  } catch (error) {
    scratches.delete(document);
    throw error;
  }
};

const EDGE = 1e-6;

function ink(shade: number): string {
  return 1.05 / (shade + 0.05) > (shade + 0.05) / 0.05 ? WHITE : BLACK;
}

/**
 * Black or white, whichever contrasts more with what `canvas` shows along a
 * caret bar from `from` to `to` (its centre line, in slide space): the median
 * luminance of opaque backing pixels whose centres are within half a slide
 * pixel of the line or whose squares it passes through (those it only touches
 * when there are none). `null` when unreadable.
 */
export function paintedCaretInk(canvas: HTMLCanvasElement | null | undefined,
  list: Pick<SlideDisplayList, 'width' | 'height'> | null | undefined,
  from: { x: number; y: number }, to: { x: number; y: number }, read: PixelRead = readThroughScratch): string | null {
  if (!canvas || !list || !list.width || !list.height || !canvas.width || !canvas.height) return null;
  const sx = canvas.width / list.width;
  const sy = canvas.height / list.height;
  const [ax, ay, bx, by] = [from.x * sx, from.y * sy, to.x * sx, to.y * sy];
  const pixelDx = bx - ax;
  const pixelDy = by - ay;
  const pixelLength = Math.hypot(pixelDx, pixelDy);
  const squareLimit = 0.5 * (Math.abs(pixelDx) + Math.abs(pixelDy));
  const pixelReach = pixelLength ? squareLimit / pixelLength : 0.5;
  const reachX = Math.max(pixelReach, 0.5 * sx);
  const reachY = Math.max(pixelReach, 0.5 * sy);
  const minX = Math.min(ax, bx);
  const minY = Math.min(ay, by);
  const maxX = Math.max(ax, bx);
  const maxY = Math.max(ay, by);
  const left = Math.max(0, Math.floor(minX - reachX));
  const top = Math.max(0, Math.floor(minY - reachY));
  const width = Math.min(canvas.width, Math.ceil(maxX + reachX)) - left;
  const height = Math.min(canvas.height, Math.ceil(maxY + reachY)) - top;
  if (!(width > 0 && height > 0)) return null;
  let data: Uint8ClampedArray | undefined;
  try { data = read(canvas, left, top, width, height); } catch { return null; }
  if (!data || data.length < width * height * 4) return null;
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const lengthSquared = dx * dx + dy * dy;
  const shades: number[] = [];
  const touched: number[] = [];
  for (let row = 0; row < height; row += 1) {
    for (let column = 0; column < width; column += 1) {
      const index = (row * width + column) * 4;
      if (data[index + 3] !== 255) continue;
      const cx = left + column + 0.5;
      const cy = top + row + 0.5;
      const box = Math.max(minX - cx, cx - maxX, minY - cy, cy - maxY) - 0.5;
      const cross = Math.abs((cx - ax) * pixelDy - (cy - ay) * pixelDx) - squareLimit;
      let into: number[] | null = box < -EDGE && cross < -EDGE * pixelLength ? shades : null;
      if (!into) {
        const x = (cx - ax) / sx;
        const y = (cy - ay) / sy;
        const t = lengthSquared ? Math.min(1, Math.max(0, (x * dx + y * dy) / lengthSquared)) : 0;
        into = Math.hypot(x - t * dx, y - t * dy) <= 0.5 ? shades : box <= EDGE && cross <= EDGE * pixelLength ? touched : null;
      }
      into?.push(luminance([data[index]!, data[index + 1]!, data[index + 2]!]));
    }
  }
  if (!shades.length) shades.push(...touched);
  if (!shades.length) return null;
  shades.sort((a, b) => a - b);
  return ink(shades[shades.length >> 1]!);
}

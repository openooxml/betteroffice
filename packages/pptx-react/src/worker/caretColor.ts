import type { GeometryPathCommand, Paint, PositionedTextLine, SlideDisplayList, SlidePrimitive } from '@betteroffice/pptx';
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

/** `estimate` is the best guess at the painted colour, `null` over an image. */
interface Fill { color: Rgb; solid: boolean; estimate: Rgb | null }

function fillOf(paint: Paint | undefined, under: Fill): Fill | null {
  const color = paintColor(paint, under.color);
  if (!color || !paint) return null;
  const opaque = paint.kind === 'solid' && (parseRgba(paint.color)?.[3] ?? 0) >= 1;
  const estimate = paint.kind !== 'solid' ? null : under.estimate ? over(paint.color, under.estimate) : opaque ? color : null;
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
    if (primitive.kind === 'shape' && upright && primitive.path.length &&
      !pathCovers(primitive.path, (x - box.x) / box.width, (y - box.y) / box.height, primitive.evenOdd)) continue;
    const exact = primitive.kind === 'shape' && primitive.geometry === 'rect' && upright;
    fill = paint && exact ? paint : { color: fill.color, solid: false, estimate: upright ? paint?.estimate ?? null : null };
  }
  return fill;
}

/** Whether a path in unit box coordinates covers `(u, v)`, curves flattened. */
function pathCovers(path: readonly GeometryPathCommand[], u: number, v: number, evenOdd = false): boolean {
  let winding = 0;
  let crossings = 0;
  let start: [number, number] = [0, 0];
  let at: [number, number] = [0, 0];
  const edge = (to: [number, number]) => {
    const [x0, y0] = at;
    const [x1, y1] = to;
    if ((y0 <= v) !== (y1 <= v) && u < x0 + (v - y0) / (y1 - y0) * (x1 - x0)) {
      crossings += 1;
      winding += y1 > y0 ? 1 : -1;
    }
    at = to;
  };
  const curve = (point: (t: number) => [number, number]) => {
    for (let step = 1; step <= 16; step += 1) edge(point(step / 16));
  };
  for (const command of path) {
    if (command.type === 'move') {
      edge(start);
      start = at = [command.x, command.y];
    } else if (command.type === 'line') edge([command.x, command.y]);
    else if (command.type === 'quad') {
      const [x0, y0] = at;
      curve((t) => [(1 - t) ** 2 * x0 + 2 * (1 - t) * t * command.cpx + t * t * command.x,
        (1 - t) ** 2 * y0 + 2 * (1 - t) * t * command.cpy + t * t * command.y]);
    } else if (command.type === 'cubic') {
      const [x0, y0] = at;
      curve((t) => [(1 - t) ** 3 * x0 + 3 * (1 - t) ** 2 * t * command.cp1x + 3 * (1 - t) * t * t * command.cp2x + t ** 3 * command.x,
        (1 - t) ** 3 * y0 + 3 * (1 - t) ** 2 * t * command.cp1y + 3 * (1 - t) * t * t * command.cp2y + t ** 3 * command.y]);
    } else edge(start);
  }
  edge(start);
  return evenOdd ? crossings % 2 === 1 : winding !== 0;
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
 * against the fill behind it, or, over a gradient or an image, black for dark
 * text and white for light text.
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
  const bar = under.estimate ? contrasting(under.estimate) : text ? matching(text) : fallback;
  if (!under.solid) {
    const color = text ? textColor! : fallback;
    const opposite = (text ? matching(text) : fallback) === WHITE ? BLACK : WHITE;
    return { color, caret: bar, background: opposite };
  }
  const background = hex(fill);
  if (!text || contrast(text, fill) < 2) return { color: fallback, caret: bar, background };
  return { color: textColor!, caret: bar, background };
}

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

interface Fill { color: Rgb; solid: boolean }

function fillOf(paint: Paint | undefined, under: Fill): Fill | null {
  const color = paintColor(paint, under.color);
  if (!color || !paint) return null;
  const opaque = paint.kind === 'solid' && (parseRgba(paint.color)?.[3] ?? 0) >= 1;
  return { color, solid: paint.kind === 'solid' && (opaque || under.solid) };
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
    const exact = primitive.kind === 'shape' && primitive.geometry === 'rect' &&
      transforms.every(({ transform }) => !transform?.rotationDeg && !transform?.flipH && !transform?.flipV);
    fill = paint && exact ? paint : { color: fill.color, solid: false };
  }
  return fill;
}

/**
 * Caret colours at a slide-space caret rect: the colour of the text at the
 * caret (its run, or the insertion colour on an empty line), as long as it
 * reads against a known solid fill; else black or white by that fill. Over a
 * gradient, an image or a shape whose painted area is not its box, the fill is
 * unknown: the text colour is used as is and `halo` gives the caret a
 * contrasting outline, so it shows on whatever is painted there.
 */
export function pptxCaretColors(list: SlideDisplayList | null | undefined, line: PositionedTextLine | null | undefined,
  position: number | null | undefined, caret: CommandRect | null | undefined,
  insertion?: string | null): { color: string; background: string; halo?: string } {
  let under: Fill = fillOf(list?.background, { color: [255, 255, 255], solid: true }) ?? { color: [255, 255, 255], solid: true };
  if (list && caret) under = fillAt(list.primitives, [], caret.x + caret.width / 2, caret.y + caret.height / 2, under);
  const fill = under.color;
  const runs = line?.runs ?? [];
  const run = position === null || position === undefined ? undefined
    : runs.find((candidate) => candidate.start < position && position <= candidate.end)
      ?? runs.find((candidate) => candidate.start <= position && position <= candidate.end);
  const textColor = run?.color ?? insertion ?? undefined;
  const text = parseColor(textColor);
  const fallback = luminance(fill) < 0.18 ? WHITE : '#000000';
  if (!under.solid) {
    const color = text ? textColor! : fallback;
    const opposite = luminance(text ?? parseColor(color)!) < 0.18 ? WHITE : '#000000';
    return { color, background: opposite, halo: opposite };
  }
  const background = hex(fill);
  if (text && contrast(text, fill) >= 2) return { color: textColor!, background };
  return { color: fallback, background };
}

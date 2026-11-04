import type { DisplayList, Rect } from '@betteroffice/xlsx';

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

function hex(rgb: Rgb): string {
  return '#' + rgb.map((channel) => channel.toString(16).padStart(2, '0')).join('');
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

function inside(box: Rect, x: number, y: number): boolean {
  return x >= box.x && x <= box.x + box.w && y >= box.y && y <= box.y + box.h;
}

/**
 * Colours for the in-cell editor over `cell`: the painted cell fill as its
 * background and the cell's text colour for text and caret, as long as that
 * reads against the fill; else black or white by the fill.
 */
export function cellEditorColors(list: DisplayList | null | undefined, cell: Rect | null): { color: string; background: string } {
  let background = WHITE;
  let fill = parseColor(WHITE)!;
  let text: string | undefined;
  if (list && cell) {
    const x = cell.x + cell.w / 2;
    const y = cell.y + cell.h / 2;
    for (const command of list.commands) {
      if (command.op === 'fillRect') {
        const rgba = parseRgba(command.color);
        if (rgba && rgba[3] >= 0.5 && inside(command, x, y) && (!command.clip || inside(command.clip, x, y))) {
          fill = [0, 1, 2].map((channel) => Math.round(rgba[channel]! * rgba[3] + fill[channel]! * (1 - rgba[3]))) as unknown as Rgb;
          background = rgba[3] < 1 ? hex(fill) : command.color;
        }
      } else if (command.op === 'text' && !command.ghost && !command.chart) {
        if (command.clip ? inside(command.clip, x, y) : inside(cell, command.x, command.y)) text = command.color;
      }
    }
  }
  const parsed = parseColor(text);
  if (text && parsed && contrast(parsed, fill) >= 2) return { color: text, background };
  return { color: luminance(fill) < 0.18 ? WHITE : '#000000', background };
}

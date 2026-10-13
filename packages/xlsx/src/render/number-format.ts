import type { NumberFormatLayout, TextAlign, TextCmd } from '../display-list/types';

/** Paints one run of `cmd` anchored at `x` by `align`, which the context's textAlign matches. */
export type PaintRun = (
  ctx: CanvasRenderingContext2D,
  cmd: TextCmd,
  text: string,
  x: number,
  align: TextAlign
) => void;

/** Lays out `cmd.text` by its parts, painting each run left-aligned. */
export function paintNumberFormat(
  ctx: CanvasRenderingContext2D,
  cmd: TextCmd,
  layout: NumberFormatLayout,
  paintRun: PaintRun
): void {
  const text = cmd.text;
  const align = cmd.align ?? 'left';
  const parts = layout.parts;
  let cursor = 0;
  let fixed = 0;
  let hasFill = false;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const index = utf16Index(text, part.offset);
    fixed += measureRange(ctx, text, cursor, index);
    if (part.kind === 'padding') {
      fixed += ctx.measureText(part.character).width;
      cursor = index + 1;
    } else {
      hasFill = true;
      cursor = index;
    }
  }
  fixed += measureRange(ctx, text, cursor, text.length);
  if (layout.overflow && fixed > layout.width) {
    const fitted = layout.general === undefined ? undefined : fitGeneral(layout.general, layout.width, ctx);
    const result = fitted ?? hashes(layout.width, ctx);
    paintRun(ctx, cmd, result, layout.left + offset(align, layout.width, ctx.measureText(result).width), 'left');
    return;
  }
  const extra = hasFill ? Math.max(0, layout.width - fixed) : 0;
  let x = layout.left + offset(align, layout.width, fixed + extra);
  cursor = 0;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const index = utf16Index(text, part.offset);
    if (index > cursor) {
      const literal = text.slice(cursor, index);
      paintRun(ctx, cmd, literal, x, 'left');
      x += ctx.measureText(literal).width;
    }
    if (part.kind === 'padding') {
      x += ctx.measureText(part.character).width;
      cursor = index + 1;
    } else {
      const advance = ctx.measureText(part.character).width;
      if (advance > 0 && extra >= advance) {
        paintRun(ctx, cmd, part.character.repeat(Math.floor(extra / advance)), x, 'left');
      }
      x += extra;
      cursor = index;
    }
  }
  if (cursor < text.length) paintRun(ctx, cmd, cursor === 0 ? text : text.slice(cursor), x, 'left');
}

/** UTF-16 index of a UTF-8 byte offset into `text`. */
function utf16Index(text: string, offset: number): number {
  let index = 0;
  for (let bytes = 0; bytes < offset && index < text.length; ) {
    const point = text.codePointAt(index)!;
    bytes += point < 0x80 ? 1 : point < 0x800 ? 2 : point < 0x10000 ? 3 : 4;
    index += point > 0xffff ? 2 : 1;
  }
  return index;
}

function measureRange(ctx: CanvasRenderingContext2D, text: string, start: number, end: number): number {
  if (end <= start) return 0;
  return ctx.measureText(start === 0 && end === text.length ? text : text.slice(start, end)).width;
}

function offset(align: TextAlign, width: number, used: number): number {
  return align === 'right' ? width - used : align === 'center' ? (width - used) / 2 : 0;
}

function hashes(width: number, ctx: CanvasRenderingContext2D): string {
  const advance = ctx.measureText('#').width;
  return advance > 0 && width > 0 ? '#'.repeat(Math.max(1, Math.floor(width / advance))) : '';
}

function fitGeneral(value: number, width: number, ctx: CanvasRenderingContext2D): string | undefined {
  if (!Number.isFinite(value) || value === 0) return undefined;
  const exponent = Math.floor(Math.log10(Math.abs(value)));
  const trim = (text: string) => text.includes('.') ? text.replace(/0+$/, '').replace(/\.$/, '') : text;
  if (exponent >= -4 && exponent < 11) {
    for (let decimals = Math.min(15, Math.max(0, 10 - exponent)); decimals >= 0; decimals--) {
      const factor = 10 ** decimals;
      const rounded = Math.sign(value) * Math.round(Math.abs(value) * factor) / factor;
      const candidate = trim(rounded.toFixed(decimals));
      if (Number(candidate) !== 0 && ctx.measureText(candidate).width <= width) return candidate;
    }
  }
  for (let decimals = 5; decimals >= 0; decimals--) {
    const [mantissa, exponent] = value.toExponential(decimals).split('e');
    const power = Number(exponent);
    const candidate = `${trim(mantissa)}E${power < 0 ? '-' : '+'}${Math.abs(power).toString().padStart(2, '0')}`;
    if (ctx.measureText(candidate).width <= width) return candidate;
  }
  return undefined;
}

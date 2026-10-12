import type { NumberFormatLayout, TextAlign } from '../display-list/types';

export function paintNumberFormat(
  text: string,
  layout: NumberFormatLayout,
  align: TextAlign,
  measure: (text: string) => number,
  paint: (text: string, x: number) => void
): void {
  let index = 0;
  let bytes = 0;
  const parts = layout.parts.map((part) => {
    while (bytes < part.offset && index < text.length) {
      const point = text.codePointAt(index)!;
      bytes += point < 0x80 ? 1 : point < 0x800 ? 2 : point < 0x10000 ? 3 : 4;
      index += point > 0xffff ? 2 : 1;
    }
    return { ...part, index };
  });
  let cursor = 0;
  let fixed = 0;
  let hasFill = false;
  for (const part of parts) {
    fixed += measure(text.slice(cursor, part.index));
    if (part.kind === 'padding') {
      fixed += measure(part.character);
      cursor = part.index + 1;
    } else {
      hasFill = true;
      cursor = part.index;
    }
  }
  fixed += measure(text.slice(cursor));
  const extra = hasFill ? Math.max(0, layout.width - fixed) : 0;
  const width = fixed + extra;
  let x = layout.left + (align === 'right' ? layout.width - width : align === 'center' ? (layout.width - width) / 2 : 0);
  cursor = 0;
  for (const part of parts) {
    const literal = text.slice(cursor, part.index);
    if (literal) {
      paint(literal, x);
      x += measure(literal);
    }
    if (part.kind === 'padding') {
      x += measure(part.character);
      cursor = part.index + 1;
    } else {
      const advance = measure(part.character);
      if (advance > 0 && extra >= advance) {
        paint(part.character.repeat(Math.floor(extra / advance)), x);
      }
      x += extra;
      cursor = part.index;
    }
  }
  if (cursor < text.length) paint(text.slice(cursor), x);
}

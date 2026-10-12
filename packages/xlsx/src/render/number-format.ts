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
  if (layout.overflow && fixed > layout.width) {
    const fitted = layout.general === undefined ? undefined : fitGeneral(layout.general, layout.width, measure);
    const advance = measure('#');
    const hashes = advance > 0 && layout.width > 0 ? '#'.repeat(Math.max(1, Math.floor(layout.width / advance))) : '';
    const result = fitted ?? hashes;
    const width = measure(result);
    const x = layout.left + (align === 'right' ? layout.width - width : align === 'center' ? (layout.width - width) / 2 : 0);
    paint(result, x);
    return;
  }
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

function fitGeneral(value: number, width: number, measure: (text: string) => number): string | undefined {
  if (!Number.isFinite(value) || value === 0) return undefined;
  const exponent = Math.floor(Math.log10(Math.abs(value)));
  const trim = (text: string) => text.includes('.') ? text.replace(/0+$/, '').replace(/\.$/, '') : text;
  if (exponent >= -4 && exponent < 11) {
    for (let decimals = Math.min(15, Math.max(0, 10 - exponent)); decimals >= 0; decimals--) {
      const factor = 10 ** decimals;
      const rounded = Math.sign(value) * Math.round(Math.abs(value) * factor) / factor;
      const candidate = trim(rounded.toFixed(decimals));
      if (Number(candidate) !== 0 && measure(candidate) <= width) return candidate;
    }
  }
  for (let decimals = 5; decimals >= 0; decimals--) {
    const [mantissa, exponent] = value.toExponential(decimals).split('e');
    const power = Number(exponent);
    const candidate = `${trim(mantissa)}E${power < 0 ? '-' : '+'}${Math.abs(power).toString().padStart(2, '0')}`;
    if (measure(candidate) <= width) return candidate;
  }
  return undefined;
}

import type { DocxEchoGlyphSet } from './residentCommandProtocol';

/** The characters a typing echo can place: printable ASCII and common typographic punctuation. */
export const DOCX_ECHO_ALPHABET = Array.from({ length: 95 }, (_, index) => String.fromCharCode(32 + index)).join('') +
  '\u2018\u2019\u201c\u201d\u2013\u2014\u2026';

/** @internal A glyph set with its lookups. */
export interface EchoGlyphIndex {
  set: DocxEchoGlyphSet;
  /** The f32 scale exactly; JSON carries its shortest decimal. */
  horizontal: number;
  /** Pixel advance of a unit count, widened from f32 as layout widens it. */
  advance(units: number): number;
  chars: ReadonlyMap<string, number>;
  kerning: ReadonlyMap<number, number>;
  unsafe: ReadonlySet<number>;
}

/** @internal */
export interface EchoGlyphPlacement {
  glyphIds: number[];
  xs: number[];
  advances: number[];
  end: number;
}

/** @internal */
export function indexEchoGlyphSet(set: DocxEchoGlyphSet): EchoGlyphIndex {
  const count = set.glyphIds.length;
  const chars = new Map<string, number>();
  [...set.alphabet].forEach((char, index) => { if (set.glyphIds[index]) chars.set(char, index); });
  const horizontal = Math.fround(set.horizontal);
  const advances = new Map<number, number>();
  return {
    set, chars, horizontal,
    advance(units) {
      let value = advances.get(units);
      if (value === undefined) advances.set(units, value = widenF32(Math.fround(units * horizontal)));
      return value;
    },
    kerning: new Map(set.kerning.map(([left, right, units]) => [left * count + right, units])),
    unsafe: new Set(set.unsafePairs.map(([left, right]) => left * count + right)),
  };
}

/** The f64 layout widens an f32 width to: the shortest decimal that still names the f32. */
export function widenF32(value: number): number {
  for (let digits = 1; digits <= 9; digits += 1) {
    const widened = Number(value.toPrecision(digits));
    if (Math.fround(widened) === value) return widened;
  }
  return value;
}

/** The engine's display-list rounding of pen positions (Rust `round3`). */
export function echoRound3(value: number): number {
  const scaled = value * 1000;
  const rounded = Math.sign(scaled) * Math.round(Math.abs(scaled)) / 1000;
  return rounded === 0 ? 0 : rounded;
}

/**
 * Places `text` as the engine lays out a cluster-positioned line from pen `x`: each character's
 * advance is its shaped units plus the pair adjustment with the character after it, scaled in f32.
 * Undefined when a character or a pair is outside the glyph set.
 */
export function placeEchoGlyphs(glyphs: EchoGlyphIndex, text: string, next: string, x: number): EchoGlyphPlacement | undefined {
  const count = glyphs.set.glyphIds.length;
  const indices: number[] = [];
  for (const char of text) {
    const index = glyphs.chars.get(char);
    if (index === undefined) return undefined;
    indices.push(index);
  }
  const following = next ? glyphs.chars.get(next) : undefined;
  if (next && following === undefined) return undefined;
  const placement: EchoGlyphPlacement = { glyphIds: [], xs: [], advances: [], end: x };
  let pen = x;
  for (let position = 0; position < indices.length; position += 1) {
    const left = indices[position]!;
    const right = position + 1 < indices.length ? indices[position + 1] : following;
    if (right !== undefined && glyphs.unsafe.has(left * count + right)) return undefined;
    const units = glyphs.set.advances[left]! + (right === undefined ? 0 : glyphs.kerning.get(left * count + right) ?? 0);
    const advance = glyphs.advance(units);
    if (!Number.isFinite(advance) || advance < 0) return undefined;
    placement.glyphIds.push(glyphs.set.glyphIds[left]!);
    placement.xs.push(echoRound3(pen));
    placement.advances.push(advance);
    pen += advance;
  }
  placement.end = pen;
  return placement;
}

/**
 * Places `text` on a line starting at `x`; with `center`, the pen starts where the engine centers
 * the line's width in that space, at half the free width and never left of `x`.
 */
export function placeEchoLine(
  glyphs: EchoGlyphIndex, text: string, next: string, x: number, center?: number
): EchoGlyphPlacement | undefined {
  if (center === undefined) return placeEchoGlyphs(glyphs, text, next, x);
  const measured = placeEchoGlyphs(glyphs, text, next, 0);
  if (!measured) return undefined;
  let width = 0;
  for (const advance of measured.advances) width += advance;
  return placeEchoGlyphs(glyphs, text, next, x + Math.max(0, (center - width) / 2));
}

/** Width of the leading letters and digits, which no line break can split. */
export function leadingWordWidth(text: string, advances: readonly number[]): number {
  let width = 0;
  let position = 0;
  for (const char of text) {
    if (!/[A-Za-z0-9]/.test(char)) break;
    width += advances[position] ?? 0;
    position += 1;
  }
  return width;
}

/** @internal Ink extent of a glyph outline in font units; undefined for a blank glyph. */
export function outlineBounds(json: string): { upem: number; xMin: number; xMax: number; yMin: number; yMax: number } | undefined {
  const outline = JSON.parse(json) as { upem: number; cmds: Array<Record<string, number | string>> };
  let xMin = Infinity;
  let xMax = -Infinity;
  let yMin = Infinity;
  let yMax = -Infinity;
  for (const command of outline.cmds) {
    for (const [key, value] of Object.entries(command)) {
      if (typeof value !== 'number') continue;
      if (key.endsWith('x')) { xMin = Math.min(xMin, value); xMax = Math.max(xMax, value); }
      if (key.endsWith('y')) { yMin = Math.min(yMin, value); yMax = Math.max(yMax, value); }
    }
  }
  return xMin <= xMax ? { upem: outline.upem, xMin, xMax, yMin, yMax } : undefined;
}

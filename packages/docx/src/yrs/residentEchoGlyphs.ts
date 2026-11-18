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

/**
 * Width of at most the least a line break moves of `text`: its leading letters and digits, which
 * no line break splits, or else its first character (East Asian text breaks between any two), or
 * nothing before a space, which hangs past the line above.
 */
export function leadingWordWidth(text: string, advances: readonly number[]): number {
  let width = 0;
  let position = 0;
  for (const char of text) {
    if (!/[A-Za-z0-9]/.test(char)) return position === 0 && !/\s/.test(char) ? advances[0] ?? 0 : width;
    width += advances[position] ?? 0;
    position += 1;
  }
  return width;
}

/** The compatibility jamo (U+3131–U+318E) that Korean input types. */
export const DOCX_ECHO_JAMO = Array.from({ length: 94 }, (_, index) => String.fromCharCode(0x3131 + index)).join('');

/** @internal How a line that mixes fonts, or spaces its letters, advances each character. */
export interface EchoLineFaces {
  /** Glyph set of each face. */
  sets: readonly EchoGlyphIndex[];
  /** Face of each code point of the placed text. */
  charFaces: readonly number[];
  /** Face of the character after the text; pairs only kern within one face. */
  nextFace?: number;
  /** Extra advance after every character, an f32 in px. */
  letterSpacing?: number;
  /** East Asian auto-spacing turned off: bit 1 against Latin letters, bit 2 against digits. */
  autoSpaceOff?: number;
  /** The f32 font size in px measurement sizes auto-spacing with. */
  size: number;
}

/** Auto-spacing class of a character, as `ooxml_text::auto_space` classifies it. */
function spacingClass(char: string): 'eastAsian' | 'latin' | 'numeric' | undefined {
  const code = char.codePointAt(0)!;
  if (code >= 0x3040 && code <= 0x30ff || code >= 0x3100 && code <= 0x312f || code >= 0x31a0 && code <= 0x31bf ||
    code >= 0x3190 && code <= 0x319f || code >= 0x3400 && code <= 0x4dbf || code >= 0x4e00 && code <= 0x9fff ||
    code >= 0xf900 && code <= 0xfaff || code >= 0xa000 && code <= 0xa4cf || code >= 0xac00 && code <= 0xd7af ||
    code >= 0x1100 && code <= 0x11ff || code >= 0x3130 && code <= 0x318f || code >= 0xff01 && code <= 0xff60 ||
    code >= 0xffe0 && code <= 0xffe6 || code >= 0x20000 && code <= 0x2fa1f) return 'eastAsian';
  if (code >= 0x2000) return undefined;
  if (code >= 0x30 && code <= 0x39) return 'numeric';
  return /\p{Alphabetic}/u.test(char) ? 'latin' : undefined;
}

/** Extra advance measurement adds to `left` before `right` (`AutoSpace::extra_px`), an f32. */
export function echoAutoSpace(left: string, right: string, size: number, off = 0): number {
  const pair = [spacingClass(left), spacingClass(right)];
  if (!pair.includes('eastAsian')) return 0;
  const other = pair[0] === 'eastAsian' ? pair[1] : pair[0];
  const spaced = other === 'latin' ? !(off & 1) : other === 'numeric' ? !(off & 2) : false;
  return spaced ? Math.fround(0.25 * size) : 0;
}

/**
 * Places `text` as the engine lays out a cluster-positioned line that mixes fonts: each character
 * advances by its own face's shaped units (and the pair adjustment with the next character of the
 * same face), scaled in f32, widened by measurement's East Asian auto-spacing and letter spacing
 * in f32. Undefined when a character or a pair is outside its face's glyph set.
 */
export function placeEchoFaceGlyphs(
  line: EchoLineFaces, text: string, next: string, x: number
): EchoGlyphPlacement | undefined {
  const chars = [...text];
  if (chars.length !== line.charFaces.length) return undefined;
  const indices: number[] = [];
  for (let position = 0; position < chars.length; position += 1) {
    const index = line.sets[line.charFaces[position]!]?.chars.get(chars[position]!);
    if (index === undefined) return undefined;
    indices.push(index);
  }
  const nextSet = next && line.nextFace !== undefined ? line.sets[line.nextFace] : undefined;
  const following = nextSet?.chars.get(next);
  const spacing = Math.fround(line.letterSpacing ?? 0);
  const placement: EchoGlyphPlacement = { glyphIds: [], xs: [], advances: [], end: x };
  let pen = x;
  for (let position = 0; position < chars.length; position += 1) {
    const face = line.charFaces[position]!;
    const glyphs = line.sets[face]!;
    const count = glyphs.set.glyphIds.length;
    const left = indices[position]!;
    const right = position + 1 < chars.length
      ? line.charFaces[position + 1] === face ? indices[position + 1] : undefined
      : line.nextFace === face ? following : undefined;
    if (right !== undefined && glyphs.unsafe.has(left * count + right)) return undefined;
    const units = glyphs.set.advances[left]! + (right === undefined ? 0 : glyphs.kerning.get(left * count + right) ?? 0);
    let advance = Math.fround(units * glyphs.horizontal);
    const after = position + 1 < chars.length ? chars[position + 1]! : next;
    const extra = after ? echoAutoSpace(chars[position]!, after, line.size, line.autoSpaceOff) : 0;
    if (extra) advance = Math.fround(advance + extra);
    if (spacing) advance = Math.fround(advance + spacing);
    const widened = widenF32(advance);
    if (!Number.isFinite(widened)) return undefined;
    placement.glyphIds.push(glyphs.set.glyphIds[left]!);
    placement.xs.push(echoRound3(pen));
    placement.advances.push(widened);
    pen += widened;
  }
  placement.end = pen;
  return placement;
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

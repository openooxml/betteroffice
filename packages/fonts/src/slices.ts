/**
 * Sliced CFF faces (`crates/betteroffice-fonts/src/slices.rs`, cut by
 * `scripts/fonts/cjk-slices.py`): a skeleton with every table of the face but
 * the CFF CharStrings INDEX, and chunks of `(glyph id, charstring)` pairs.
 * Appending a CharStrings INDEX built from any set of chunks reassembles a
 * complete face.
 */

const MAGIC = 0x424f534b;
const FORMAT = 1;
const HEADER = 32;
const RUN = 12;
const ENDCHAR = new Uint8Array([0x0e]);

const UNSUPPORTED = 'Not a supported font slice skeleton';
const TRUNCATED = 'Truncated font slice chunk';

/** Chunk id of a glyph or codepoint no chunk holds. */
export const NO_CHUNK = 0xffff;

export interface SliceSkeleton {
  readonly numGlyphs: number;
  /** The chunk holding `glyph`'s outline, or `undefined`. */
  chunkOfGlyph(glyph: number): number | undefined;
  /** The chunk a codepoint's glyphs are cut into, or `undefined`. */
  chunkOfCodepoint(codepoint: number): number | undefined;
  /** The face with `outlines` by glyph id; every other glyph is `endchar`. */
  reassemble(outlines: ReadonlyMap<number, Uint8Array>): ArrayBuffer;
}

function fits(offset: number, length: number, limit: number): boolean {
  return offset + length <= limit;
}

export function parseSkeleton(bytes: Uint8Array): SliceSkeleton {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.byteLength < HEADER || view.getUint32(0) !== MAGIC || view.getUint16(4) !== FORMAT) throw new Error(UNSUPPORTED);
  const numGlyphs = view.getUint32(8);
  const cffRecord = view.getUint32(12);
  const cffOffset = view.getUint32(16);
  const headOffset = view.getUint32(20);
  const prefixLength = view.getUint32(24);
  const runCount = view.getUint32(28);
  const gidEnd = HEADER + numGlyphs * 2;
  const runsEnd = gidEnd + runCount * RUN;
  if (prefixLength > bytes.byteLength || runsEnd !== bytes.byteLength - prefixLength || numGlyphs > 0xffff
    || !fits(cffRecord, 16, prefixLength) || !fits(cffOffset, 0, prefixLength)
    || !fits(headOffset, 12, prefixLength)) {
    throw new Error(UNSUPPORTED);
  }
  const prefix = bytes.subarray(runsEnd);
  return {
    numGlyphs,
    chunkOfGlyph(glyph) {
      if (!Number.isInteger(glyph) || glyph < 0 || glyph >= numGlyphs) return undefined;
      const chunk = view.getUint16(HEADER + glyph * 2);
      return chunk === NO_CHUNK ? undefined : chunk;
    },
    chunkOfCodepoint(codepoint) {
      let low = 0;
      let high = runCount;
      while (low < high) {
        const middle = (low + high) >>> 1;
        const at = gidEnd + middle * RUN;
        const start = view.getUint32(at);
        if (codepoint < start) high = middle;
        else if (codepoint - start >= view.getUint32(at + 4)) low = middle + 1;
        else return view.getUint16(at + 8);
      }
      return undefined;
    },
    reassemble(outlines) {
      let dataLength = 0;
      for (let glyph = 0; glyph < numGlyphs; glyph += 1) {
        dataLength += (outlines.get(glyph) ?? ENDCHAR).byteLength;
      }
      const indexLength = 3 + (numGlyphs + 1) * 4 + dataLength;
      const length = prefix.byteLength + indexLength;
      const font = new Uint8Array(length + ((4 - (length % 4)) % 4));
      font.set(prefix);
      const out = new DataView(font.buffer);
      let at = prefix.byteLength;
      out.setUint16(at, numGlyphs);
      out.setUint8(at + 2, 4);
      at += 3;
      let offset = 1;
      out.setUint32(at, offset);
      at += 4;
      for (let glyph = 0; glyph < numGlyphs; glyph += 1) {
        offset += (outlines.get(glyph) ?? ENDCHAR).byteLength;
        out.setUint32(at, offset);
        at += 4;
      }
      for (let glyph = 0; glyph < numGlyphs; glyph += 1) {
        const outline = outlines.get(glyph) ?? ENDCHAR;
        font.set(outline, at);
        at += outline.byteLength;
      }
      out.setUint32(cffRecord + 4, checksum(font, cffOffset, length));
      out.setUint32(cffRecord + 12, length - cffOffset);
      out.setUint32(headOffset + 8, 0);
      out.setUint32(headOffset + 8, (0xb1b0afba - checksum(font, 0, font.byteLength)) >>> 0);
      return font.buffer;
    },
  };
}

/** `(glyph id, charstring)` pairs of a chunk file; the charstrings view `bytes`. */
export function chunkEntries(bytes: Uint8Array): Array<[number, Uint8Array]> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.byteLength < 4) throw new Error(TRUNCATED);
  const count = view.getUint32(0);
  const entries: Array<[number, Uint8Array]> = [];
  let at = 4;
  for (let index = 0; index < count; index += 1) {
    if (at + 8 > bytes.byteLength) throw new Error(TRUNCATED);
    const glyph = view.getUint32(at);
    const length = view.getUint32(at + 4);
    const end = at + 8 + length;
    if (end > bytes.byteLength) throw new Error(TRUNCATED);
    entries.push([glyph, bytes.subarray(at + 8, end)]);
    at = end;
  }
  if (at !== bytes.byteLength) throw new Error('Malformed font slice chunk');
  return entries;
}

function checksum(bytes: Uint8Array, start: number, end: number): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let sum = 0;
  let at = start;
  for (; at + 4 <= end; at += 4) sum = (sum + view.getUint32(at)) >>> 0;
  if (at < end) {
    let word = 0;
    for (let shift = 24; at < end; at += 1, shift -= 8) word |= bytes[at]! << shift;
    sum = (sum + (word >>> 0)) >>> 0;
  }
  return sum;
}

const MARK_CLUSTER = /\P{M}\p{M}+/gu;
const MOST_MARKS = 8;
const RIGHT_TO_LEFT = /[\u0590-\u08ff\u200f\u202b\u202e\u2067\ufb1d-\ufdff\ufe70-\ufeff\u{10800}-\u{10fff}\u{1e800}-\u{1efff}]/u;
const MIRRORED = /\p{Bidi_Mirrored}/u;
let mirroredCharacters: string | undefined;

/** Every character right-to-left text can be drawn with in place of another. */
function mirrored(): string {
  if (mirroredCharacters === undefined) {
    mirroredCharacters = '';
    for (let codepoint = 0x20; codepoint < 0x20000; codepoint += 1) {
      const character = String.fromCodePoint(codepoint);
      if (MIRRORED.test(character)) mirroredCharacters += character;
    }
  }
  return mirroredCharacters;
}

/** Every character a shaper can compose from `text`, including a base that takes only some of its marks. */
function compositions(text: string): string {
  let composed = text.normalize('NFC');
  for (const cluster of text.normalize('NFD').match(MARK_CLUSTER) ?? []) {
    const [base, ...marks] = [...cluster];
    for (let picked = 1; picked < 1 << Math.min(marks.length, MOST_MARKS); picked += 1) {
      composed += `${base}${marks.filter((_, index) => picked & 1 << index).join('')}`.normalize('NFC');
    }
  }
  return composed;
}

/**
 * The chunks shaping `text` can draw from. GSUB results sit in the chunks of their characters, but a shaper also
 * composes canonical sequences, decomposes characters the face lacks and mirrors right-to-left text, which reach
 * other chunks.
 */
export function textChunks(text: string, chunkOfCodepoint: (codepoint: number) => number | undefined): number[] {
  const chunks: number[] = [];
  const flipped = RIGHT_TO_LEFT.test(text) && MIRRORED.test(text) ? mirrored() : '';
  for (const character of `${text}${compositions(text)}${flipped}`) {
    const chunk = chunkOfCodepoint(character.codePointAt(0)!);
    if (chunk !== undefined) { chunks.push(chunk); continue; }
    for (const part of character.normalize('NFD')) {
      const inner = chunkOfCodepoint(part.codePointAt(0)!);
      if (inner !== undefined) chunks.push(inner);
    }
  }
  return chunks;
}

// @betteroffice/docx declaration yrs/residentEchoGlyphs.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxEchoGlyphSet } from './residentCommandProtocol';
/** The characters a typing echo can place: printable ASCII and common typographic punctuation. */
export declare const DOCX_ECHO_ALPHABET: string;
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
export declare function indexEchoGlyphSet(set: DocxEchoGlyphSet): EchoGlyphIndex;
/** The f64 layout widens an f32 width to: the shortest decimal that still names the f32. */
export declare function widenF32(value: number): number;
/** The engine's display-list rounding of pen positions (Rust `round3`). */
export declare function echoRound3(value: number): number;
/**
 * Places `text` as the engine lays out a cluster-positioned line from pen `x`: each character's
 * advance is its shaped units plus the pair adjustment with the character after it, scaled in f32.
 * Undefined when a character or a pair is outside the glyph set.
 */
export declare function placeEchoGlyphs(glyphs: EchoGlyphIndex, text: string, next: string, x: number): EchoGlyphPlacement | undefined;
/**
 * Places `text` on a line starting at `x`; with `center`, the pen starts where the engine centers
 * the line's width in that space, at half the free width and never left of `x`.
 */
export declare function placeEchoLine(glyphs: EchoGlyphIndex, text: string, next: string, x: number, center?: number): EchoGlyphPlacement | undefined;
/**
 * Width of at most the least a line break moves of `text`: its leading letters and digits, which
 * no line break splits, or else its first character (East Asian text breaks between any two), or
 * nothing before a space, which hangs past the line above.
 */
export declare function leadingWordWidth(text: string, advances: readonly number[]): number;
/** The compatibility jamo (U+3131–U+318E) that Korean input types. */
export declare const DOCX_ECHO_JAMO: string;
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
/** Extra advance measurement adds to `left` before `right` (`AutoSpace::extra_px`), an f32. */
export declare function echoAutoSpace(left: string, right: string, size: number, off?: number): number;
/**
 * Places `text` as the engine lays out a cluster-positioned line that mixes fonts: each character
 * advances by its own face's shaped units (and the pair adjustment with the next character of the
 * same face), scaled in f32, widened by measurement's East Asian auto-spacing and letter spacing
 * in f32. Undefined when a character or a pair is outside its face's glyph set, or when letter
 * spacing condenses a character to a negative advance.
 */
export declare function placeEchoFaceGlyphs(line: EchoLineFaces, text: string, next: string, x: number): EchoGlyphPlacement | undefined;
/** @internal Ink extent of a glyph outline in font units; undefined for a blank glyph. */
export declare function outlineBounds(json: string): {
    upem: number;
    xMin: number;
    xMax: number;
    yMin: number;
    yMax: number;
} | undefined;

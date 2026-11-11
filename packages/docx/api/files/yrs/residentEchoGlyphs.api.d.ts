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
/** Width of the leading letters and digits, which no line break can split. */
export declare function leadingWordWidth(text: string, advances: readonly number[]): number;
/** @internal Ink extent of a glyph outline in font units; undefined for a blank glyph. */
export declare function outlineBounds(json: string): {
    upem: number;
    xMin: number;
    xMax: number;
    yMin: number;
    yMax: number;
} | undefined;

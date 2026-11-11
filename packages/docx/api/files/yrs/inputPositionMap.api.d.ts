// @betteroffice/docx declaration yrs/inputPositionMap.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc } from './index';
/** One native paragraph in the yrs-backed display-position projection. */
export interface YrsInputParagraphSpan {
    paraId: string;
    /** Number of UTF-16/inline units before the paragraph's pilcrow. */
    length: number;
    /** Block start emitted by the yrs display-position projection. */
    displayStart: number;
}
/**
 * Live converter between display-list positions and public yrs Locs.
 *
 * A native paragraph of length N occupies N+2 display positions (open token,
 * content, close token), so this map can be rebuilt directly from the live yrs
 * paragraph spans after every structural edit.
 */
export interface YrsInputPositionMap {
    story: string;
    paragraphs: readonly YrsInputParagraphSpan[];
}
export interface YrsInputParagraphSeed {
    paraId: string;
    length: number;
}
export declare function createYrsInputPositionMap(story: string, seeds: readonly YrsInputParagraphSeed[]): YrsInputPositionMap;
/** Convert a live yrs location to the position used by display-list geometry. */
export declare function yrsLocToDisplayPosition(map: YrsInputPositionMap, loc: YrsLoc): number;
/** Convert a display-list hit position to a paragraph-keyed yrs location. */
export declare function displayPositionToYrsLoc(map: YrsInputPositionMap, position: number): YrsLoc | null;
/** Compare two Locs in document order using one live map. */
export declare function compareYrsLocs(map: YrsInputPositionMap, a: YrsLoc, b: YrsLoc): number;

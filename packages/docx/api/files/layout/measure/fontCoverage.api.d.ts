// @betteroffice/docx declaration layout/measure/fontCoverage.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
export type FontCoverage = (character: string) => boolean;
export declare function charactersFromBits(bits?: string): string[];
/** Coverage answering from base Latin bit set words. */
export declare function wordsCoverage(words: Uint32Array): FontCoverage;
/** Base Latin coverage from coverage bits (hex words). */
export declare function coverageFromBits(bits: string): FontCoverage;
export declare function baseLatinAlias(character: string): boolean;
export declare function symbolFamily(family: string): boolean;
export declare function baseLatinCandidate(character: string, styled: boolean): boolean;

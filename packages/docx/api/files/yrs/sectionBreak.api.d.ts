// @betteroffice/docx declaration yrs/sectionBreak.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { SectionProperties } from '../types/document';
import type { YrsLoc, YrsSession, YrsStorySegment } from './index';
import type { ResidentPartMetadata } from './residentParts';
import type { EditSession } from './wasm/index';
export type SectionStart = NonNullable<SectionProperties['sectionStart']>;
type BreakSession = Pick<YrsSession, 'paragraphs' | 'splitParagraph' | 'setParagraphAttrs'>;
type BreakWasm = Pick<EditSession, 'section_margins_json' | 'set_section_margins' | 'story_segment_units'>;
/** Count of the block embeds (tables, block SDTs, page and column breaks) leading a paragraph. */
export declare function leadingBlockCount(segments: readonly YrsStorySegment[]): number;
/**
 * Where a page break at `at` can land as a block: past the paragraph's
 * leading blocks, splitting the paragraph first when inline content precedes `at`.
 */
export declare function pageBreakTarget(session: BreakSession, wasm: BreakWasm, at: YrsLoc): YrsLoc;
/**
 * Splits the body paragraph at `at` into a new section boundary: the first half
 * closes with `properties(closing)` and the section's closing paragraph, when
 * one follows, starts with `start`. `closed` is false when the split falls in
 * the final section, whose start lives in the part metadata.
 */
export declare function splitSection(session: BreakSession, wasm: BreakWasm, at: YrsLoc, properties: (closing: SectionProperties | undefined) => SectionProperties, start: SectionStart): {
    at: YrsLoc;
    closed: boolean;
};
/** Page and section breaks split paragraphs, which suggesting mode cannot track. */
export declare function breakRefusedWhileSuggesting(kind: string, suggesting: boolean): boolean;
/**
 * Part metadata whose final section starts with `start`: the stored metadata
 * when a session seeded it, else the opened package's save keys with each
 * non-final section keyed by the body paragraph that closed it when the package
 * was opened. Undefined when neither exists.
 */
export declare function finalSectionStartMetadata(stored: Partial<ResidentPartMetadata>, saved: string | undefined, closing: readonly string[], start: SectionStart): ResidentPartMetadata | undefined;
/** Body paragraphs that close a section, in document order. */
export declare function sectionClosingParaIds(session: Pick<YrsSession, 'paragraphs'>): string[];
export {};

// @betteroffice/docx declaration yrs/tableBoundaryDelete.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsParagraphLength, YrsStorySegment } from './index';
interface TableDeleteReader {
    paragraphSpans(story: string): readonly YrsParagraphLength[];
    paragraphSegments(story: string, indices: readonly number[]): YrsStorySegment[][];
    /**
     * Index of the first paragraph of `story` with `paraId`, in `paragraphSpans` order, or -1.
     * When given, a delete inside a paragraph reads only the paragraphs it touches, not every
     * span of the story.
     */
    paragraphUnit?(story: string, paraId: string): number;
}
/** @internal A story's paragraph spans, read one paragraph at a time when the reader can find a paragraph by id. */
export interface StoryParagraphSpans {
    /** Index of the first paragraph with `paraId`, in `paragraphSpans` order, or -1. */
    indexOf(paraId: string): number;
    /** The span of the paragraph at `index`, which must be one of the story's paragraphs. */
    at(index: number): YrsParagraphLength;
    /** How many paragraphs the story has. Reads every span of the story. */
    count(): number;
}
/** @internal */
export declare function storyParagraphSpans(reader: TableDeleteReader, story: string): StoryParagraphSpans;
/** @internal */
export declare function tableBoundaryDeleteTarget(reader: TableDeleteReader, at: YrsLoc, direction: 'backward' | 'forward'): YrsLoc | null;
export {};

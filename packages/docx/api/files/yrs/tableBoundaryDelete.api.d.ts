// @betteroffice/docx declaration yrs/tableBoundaryDelete.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsParagraphLength, YrsStorySegment } from './index';
interface TableDeleteReader {
    paragraphSpans(story: string): readonly YrsParagraphLength[];
    paragraphSegments(story: string, indices: readonly number[]): YrsStorySegment[][];
}
/** @internal */
export declare function tableBoundaryDeleteTarget(reader: TableDeleteReader, at: YrsLoc, direction: 'backward' | 'forward'): YrsLoc | null;
export {};

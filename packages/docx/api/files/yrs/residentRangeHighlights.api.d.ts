// @betteroffice/docx declaration yrs/residentRangeHighlights.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import { type DisplayListRect } from '../layout/render/displayListQueries';
import type { RetainedFrame } from '../layout/render/frameDelta';
export interface DocxRangeHighlights {
    version: string;
    layoutRevision: number;
    frameEpoch: number;
    rects: Array<Array<DisplayListRect & {
        pageId: string;
    }>>;
    unbuiltPages: number[];
}
export declare function readResidentRangeHighlights(engine: ResidentEngineSession, frame: RetainedFrame, ranges: readonly YrsStoryRange[], layoutRevision: number): Promise<DocxRangeHighlights>;

// @betteroffice/docx declaration yrs/residentReview.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsCommentInfo, YrsRevisionInfo, YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxProposalRegistry, DocxProposalSnapshot } from './proposals';
import { type ProposalGeometryMirror } from './proposalGeometry';
import { DisplayPositionIndex } from './displayPositionIndex';
import { type DisplayListQueries, type DisplayListRect } from '../layout/render/displayListQueries';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { Comment } from '../types/content';
export interface DocxCommentPatch {
    body?: unknown;
    author?: string;
    date?: string;
    done?: boolean;
    parentId?: string | null;
}
/** A comment, by the id saves and the host use, or a revision, by its session id. */
export type DocxReviewTarget = {
    kind: 'comment';
    id: number;
} | {
    kind: 'revision';
    revisionId: string;
};
/** The session id and story ranges of a review target, or null when it does not exist. */
export interface DocxReviewTargetRanges {
    id: string;
    ranges: YrsStoryRange[];
}
export interface DocxReviewSnapshot {
    version: string;
    previewVersion: number;
    layoutRevision: number;
    frameEpoch: number;
    comments: YrsCommentInfo[];
    uiComments?: Comment[];
    commentIds?: Record<string, number>;
    revisions: YrsRevisionInfo[];
    proposals: DocxProposalSnapshot;
    proposalGeometry: ProposalGeometryMirror;
    commentGeometry: Record<string, DisplayListRect[]>;
    revisionGeometry: Record<string, DisplayListRect[]>;
    proposalHighlights: Record<string, DisplayListRect[]>;
}
export declare function readReviewTarget(engine: ResidentEngineSession, target: DocxReviewTarget): DocxReviewTargetRanges | null;
export declare function readResidentReviewSnapshot(engine: ResidentEngineSession, registry: DocxProposalRegistry, frame: RetainedFrame, layoutRevision: number, reads?: {
    queries?: (frame: RetainedFrame, version: string) => DisplayListQueries | undefined;
    index?: DisplayPositionIndex;
}): Promise<DocxReviewSnapshot>;

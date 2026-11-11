// @betteroffice/docx declaration yrs/residentCommentProjection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Comment } from '../types/content';
import type { DocxEditFailureCode } from './edits';
import type { YrsCommentInfo, YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
export interface ResidentCommentRecoveryState {
    aliases: [string, number][];
    initialized: boolean;
}
export interface ResidentCommentProjection {
    comments: Comment[];
    ids: Record<string, number>;
    sessionIds: Map<number, string>;
}
/**
 * The IDs a save gives the session's comments, in synced comment order: a source comment keeps its
 * ID, and any other comment keeps its own ID when that is a free OOXML comment ID and takes the next
 * free one after the source IDs otherwise.
 */
export declare function savedCommentIds(sourceIds: readonly number[], native: readonly Pick<YrsCommentInfo, 'id'>[], previous?: ReadonlyMap<string, number>): Map<string, number>;
/**
 * The comments a save writes: the source package's, each with the fields the session wrote over its
 * seeded placeholders and without the ones it removed, then the comments added in the session.
 * With `initialized`, native resolved state replaces the source state.
 * Session comments keep the IDs `aliases` handed out, in the order this replica first projected them;
 * with `fresh`, they take {@link savedCommentIds} instead, so replicas holding the same comments save
 * the same IDs.
 */
export declare function projectResidentComments(engine: ResidentEngineSession, aliases: Map<string, number>, initialized?: boolean, fresh?: boolean): ResidentCommentProjection;
interface CommentInsertionRefusal {
    status: 'refused';
    reason: Extract<DocxEditFailureCode, 'missing-target' | 'ambiguous-target' | 'invalid-step'>;
}
export type ResidentCommentInsertionRange = YrsStoryRange | CommentInsertionRefusal;
export declare function readCommentInsertionRange(engine: ResidentEngineSession, paraId: string, search?: string): ResidentCommentInsertionRange;
export {};

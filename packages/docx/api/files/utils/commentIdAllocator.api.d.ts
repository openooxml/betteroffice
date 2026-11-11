// @betteroffice/docx declaration utils/commentIdAllocator.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Numeric OOXML comment-id allocation, independent of any editor engine. */
export declare const PENDING_COMMENT_ID = -1;
export interface CommentIdAllocator {
    next(): number;
    seedAbove(maxId: number): void;
}
export declare function createCommentIdAllocator(): CommentIdAllocator;

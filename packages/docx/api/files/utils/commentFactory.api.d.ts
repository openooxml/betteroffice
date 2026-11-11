// @betteroffice/docx declaration utils/commentFactory.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Comment } from '../types/content';
import type { CommentIdAllocator } from './commentIdAllocator';
/** Build a comment thread entry with a freshly allocated numeric OOXML id. */
export declare function createComment(allocator: CommentIdAllocator, text: string, authorName: string, parentId?: number): Comment;

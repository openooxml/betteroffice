import type { Comment } from '../types/content';
import type { CommentIdAllocator } from './commentIdAllocator';

/** Build a comment thread entry with a freshly allocated numeric OOXML id. */
export function createComment(
  allocator: CommentIdAllocator,
  text: string,
  authorName: string,
  parentId?: number
): Comment {
  return {
    id: allocator.next(),
    author: authorName,
    date: new Date().toISOString(),
    content: [
      {
        type: 'paragraph',
        formatting: {},
        content: [{ type: 'run', formatting: {}, content: [{ type: 'text', text }] }],
      },
    ],
    ...(parentId !== undefined && { parentId }),
  };
}

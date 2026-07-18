/** Numeric OOXML comment-id allocation, independent of any editor engine. */

export const PENDING_COMMENT_ID = -1;

export interface CommentIdAllocator {
  next(): number;
  seedAbove(maxId: number): void;
}

export function createCommentIdAllocator(): CommentIdAllocator {
  let nextId = 1;
  return {
    next: () => nextId++,
    seedAbove(maxId: number) {
      if (maxId >= nextId) nextId = maxId + 1;
    },
  };
}

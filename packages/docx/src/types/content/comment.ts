/**
 * Comments (`w:comment` in `comments.xml`) and the inline range markers
 * (`w:commentRangeStart`/`End`) that anchor them inside paragraphs.
 */

import type { Paragraph } from './paragraph';
import type { BlockContent } from './section';

/** Stable reviewer/color assignment shared by comments and revisions. */
export interface CommentAuthor {
  id?: string;
  name?: string;
  initials?: string;
  paletteIndex?: number;
  color?: string;
}

/** Exact model anchor for a comment range. */
export interface CommentAnchorRange {
  startBlockId?: string | number;
  startOffset?: number;
  endBlockId?: string | number;
  endOffset?: number;
  collapsed?: boolean;
  objectId?: string;
}

/**
 * A comment from `comments.xml` — the top-level entity for review
 * comments and replies. `id` matches the inline `CommentRangeStart` /
 * `CommentRangeEnd` markers that anchor it inside a paragraph; `parentId`
 * threads replies under their parent; `done` reflects Word's "Resolve"
 * state (`w15:done`).
 */
export interface Comment {
  /** Comment ID (matches commentRangeStart/End) */
  id: number;
  /** Author name */
  author: string;
  /** Author initials */
  initials?: string;
  /** Date */
  date?: string;
  /** Comment content (paragraphs) */
  content: Paragraph[];
  /** Parent comment ID (for replies) */
  parentId?: number;
  /** Whether the comment is resolved/done */
  done?: boolean;
  /** Explicit status. Undefined is derived from `done` (false/active). */
  status?: 'active' | 'resolved';
  /** Stable reviewer/person id from modern comment parts. */
  authorId?: string;
  /** Durable modern-comment id. */
  durableId?: string;
  /** Paragraph id used to join commentsExtended metadata. */
  paraId?: string;
  /** UTC timestamp from modern comment metadata. */
  dateUtc?: string;
  /** Stable author-palette index. Undefined = derive by first occurrence. */
  paletteIndex?: number;
  /** Exact range/object anchor. Undefined = use inline markers. */
  anchorRange?: CommentAnchorRange;
  /** Rich block content. Undefined = use legacy paragraph-only `content`. */
  blockContent?: BlockContent[];
}

/**
 * Comment range start marker in paragraph content
 */
export interface CommentRangeStart {
  type: 'commentRangeStart';
  id: number;
  /** Block-local offset. Undefined = marker position in source order. */
  offset?: number;
}

/**
 * Comment range end marker in paragraph content
 */
export interface CommentRangeEnd {
  type: 'commentRangeEnd';
  id: number;
  /** Block-local offset. Undefined = marker position in source order. */
  offset?: number;
}

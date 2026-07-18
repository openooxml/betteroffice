/**
 * Reply-range marker injection for serialization.
 *
 * Word / Pages / LibreOffice expect every comment in `comments.xml`
 * (including REPLY threads) to have matching `commentRangeStart` /
 * `commentRangeEnd` / `commentReference` markers in `document.xml`.
 * The PM document only stamps marks for the parent comment because
 * replies don't have their own visible range — they share the parent
 * thread's text. So before serialization we walk the body content and
 * synthesize parallel range markers for every reply.
 *
 * Two helpers, one per parent shape:
 * - `injectReplyRangeMarkers` — replies whose parent is another
 *   comment (regular threaded discussion). Finds the parent's
 *   `commentRangeStart`/`End` and adds parallel markers next to them.
 * - `injectTCReplyRangeMarkers` — replies whose parent is a tracked
 *   change (insertion/deletion). Wraps the TC content with
 *   commentRange markers.
 *
 * Pre-#... this code lived inside React's DocxEditor.tsx; Vue had no
 * equivalent and so silently lost reply markers when saving collab
 * documents. Living in core means both adapters get it for free.
 */

import type { BlockContent, Comment, ParagraphContent } from '../types/content';

/**
 * Inject `commentRangeStart`/`commentRangeEnd` for reply comments
 * that share their parent comment's text range.
 */
export function injectReplyRangeMarkers(content: BlockContent[], comments: Comment[]): void {
  const replies = comments.filter((c) => c.parentId != null);
  if (replies.length === 0) return;

  // Build parentId → reply IDs map
  const replyIdsByParent = new Map<number, number[]>();
  for (const r of replies) {
    const arr = replyIdsByParent.get(r.parentId!);
    if (arr) arr.push(r.id);
    else replyIdsByParent.set(r.parentId!, [r.id]);
  }

  function walkBlocks(blocks: BlockContent[]): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        // Skip paragraphs without any comment range markers
        if (
          !block.content.some((i) => i.type === 'commentRangeStart' || i.type === 'commentRangeEnd')
        )
          continue;
        const newItems: ParagraphContent[] = [];
        for (const item of block.content) {
          if (item.type === 'commentRangeStart') {
            newItems.push(item);
            const replyIds = replyIdsByParent.get(item.id);
            if (replyIds) {
              for (const rid of replyIds) {
                newItems.push({ type: 'commentRangeStart', id: rid });
              }
            }
          } else if (item.type === 'commentRangeEnd') {
            newItems.push(item);
            const replyIds = replyIdsByParent.get(item.id);
            if (replyIds) {
              for (const rid of replyIds) {
                newItems.push({ type: 'commentRangeEnd', id: rid });
              }
            }
          } else {
            newItems.push(item);
          }
        }
        block.content = newItems;
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walkBlocks(cell.content);
          }
        }
      }
    }
  }

  walkBlocks(content);
}

/**
 * Inject `commentRangeStart`/`commentRangeEnd` for comments whose
 * parent is a tracked-change revision (insertion/deletion). The TC
 * content nodes don't carry the comment's range, so we wrap them.
 */
export function injectTCReplyRangeMarkers(content: BlockContent[], comments: Comment[]): void {
  const commentIds = new Set(comments.map((c) => c.id));
  const tcReplies = comments.filter((c) => c.parentId != null && !commentIds.has(c.parentId));
  if (tcReplies.length === 0) return;

  const replyIdsByRevision = new Map<number, number[]>();
  for (const r of tcReplies) {
    const arr = replyIdsByRevision.get(r.parentId!);
    if (arr) arr.push(r.id);
    else replyIdsByRevision.set(r.parentId!, [r.id]);
  }

  function walkBlocks(blocks: BlockContent[]): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        const hasTC = block.content.some(
          (item) =>
            (item.type === 'insertion' || item.type === 'deletion') &&
            replyIdsByRevision.has(item.info.id)
        );
        if (!hasTC) continue;

        const newItems: ParagraphContent[] = [];
        const items = block.content;
        for (let i = 0; i < items.length; i++) {
          const item = items[i];
          if (
            (item.type === 'insertion' || item.type === 'deletion') &&
            replyIdsByRevision.has(item.info.id)
          ) {
            const replyIds = replyIdsByRevision.get(item.info.id)!;
            for (const rid of replyIds) {
              newItems.push({ type: 'commentRangeStart', id: rid });
            }
            newItems.push(item);
            // Adjacent del+ins replacement pair share author+date —
            // include the second half inside the comment range so we
            // don't break del-ins adjacency in the saved doc.
            const next = items[i + 1];
            if (
              next &&
              (next.type === 'insertion' || next.type === 'deletion') &&
              next.type !== item.type &&
              next.info.author === item.info.author &&
              next.info.date === item.info.date
            ) {
              newItems.push(next);
              i++;
            }
            for (const rid of replyIds) {
              newItems.push({ type: 'commentRangeEnd', id: rid });
            }
          } else {
            newItems.push(item);
          }
        }
        block.content = newItems;
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walkBlocks(cell.content);
          }
        }
      }
    }
  }

  walkBlocks(content);
}

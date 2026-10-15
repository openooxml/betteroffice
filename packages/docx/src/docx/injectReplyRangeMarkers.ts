/**
 * Reply-range marker injection for serialization.
 *
 * Word / Pages / LibreOffice expect every comment in `comments.xml`
 * (including REPLY threads) to have matching `commentRangeStart` /
 * `commentRangeEnd` / `commentReference` markers in `document.xml`.
 * The document only stamps marks for the parent comment because
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

import { visitTrackedControlContent } from '../utils/trackedControlContent';

import type { BlockContent, Comment, ParagraphContent } from '../types/content';

/** Every comment ID a range marker in `content` carries, at any depth. */
function rangedCommentIds(content: BlockContent[]): Set<number> {
  const ids = new Set<number>();
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const item of node) visit(item);
      return;
    }
    if (!node || typeof node !== 'object' || ArrayBuffer.isView(node)) return;
    const entry = node as { type?: unknown; id?: unknown };
    if (
      (entry.type === 'commentRangeStart' || entry.type === 'commentRangeEnd') &&
      typeof entry.id === 'number'
    ) {
      ids.add(entry.id);
    }
    for (const value of Object.values(node)) {
      if (value && typeof value === 'object') visit(value);
    }
  };
  visit(content);
  return ids;
}

/**
 * Inject `commentRangeStart`/`commentRangeEnd` for reply comments
 * that share their parent comment's text range and have none of their own.
 */
export function injectReplyRangeMarkers(content: BlockContent[], comments: Comment[]): boolean {
  const threaded = comments.filter((c) => c.parentId != null);
  if (threaded.length === 0) return false;
  const ranged = rangedCommentIds(content);
  const replies = threaded.filter((c) => !ranged.has(c.id));
  if (replies.length === 0) return false;
  let inserted = false;

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
                inserted = true;
              }
            }
          } else if (item.type === 'commentRangeEnd') {
            newItems.push(item);
            const replyIds = replyIdsByParent.get(item.id);
            if (replyIds) {
              for (const rid of replyIds) {
                newItems.push({ type: 'commentRangeEnd', id: rid });
                inserted = true;
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
  return inserted;
}

/**
 * Inject `commentRangeStart`/`commentRangeEnd` for comments whose
 * parent is a tracked-change revision (insertion/deletion) and that have
 * no range of their own. The TC content nodes don't carry the comment's
 * range, so we wrap them.
 */
export function injectTCReplyRangeMarkers(content: BlockContent[], comments: Comment[]): boolean {
  const commentIds = new Set(comments.map((c) => c.id));
  const threaded = comments.filter((c) => c.parentId != null && !commentIds.has(c.parentId));
  if (threaded.length === 0) return false;
  const ranged = rangedCommentIds(content);
  const tcReplies = threaded.filter((c) => !ranged.has(c.id));
  if (tcReplies.length === 0) return false;
  let inserted = false;

  const replyIdsByRevision = new Map<number, number[]>();
  for (const r of tcReplies) {
    const arr = replyIdsByRevision.get(r.parentId!);
    if (arr) arr.push(r.id);
    else replyIdsByRevision.set(r.parentId!, [r.id]);
  }

  const noReplies: number[] = [];
  function revisionReplies(item: ParagraphContent): number[] {
    if (item.type === 'run') return noReplies;
    const ordinary = item.type === 'insertion' || item.type === 'deletion'
      ? replyIdsByRevision.get(item.info.id) ?? noReplies
      : noReplies;
    let nested: Set<number> | undefined;
    visitTrackedControlContent(item, (node) => {
      let revisionId: number | undefined;
      if (
        node.type === 'insertion' || node.type === 'deletion' ||
        node.type === 'moveFrom' || node.type === 'moveTo'
      ) revisionId = node.info.id;
      else if (node.type === 'inlineSdt' && (
        item.type === 'insertion' || item.type === 'deletion' ||
        item.type === 'moveFrom' || item.type === 'moveTo'
      )) revisionId = item.info.id;
      if (revisionId === undefined) return;
      const ids = replyIdsByRevision.get(revisionId);
      if (!ids) return;
      nested ??= new Set(ordinary);
      for (const id of ids) nested.add(id);
    });
    return nested ? [...nested] : ordinary;
  }

  function walkBlocks(blocks: BlockContent[]): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        const hasTC = block.content.some((item) => revisionReplies(item).length > 0);
        if (!hasTC) continue;

        const newItems: ParagraphContent[] = [];
        const items = block.content;
        for (let i = 0; i < items.length; i++) {
          const item = items[i];
          const replyIds = revisionReplies(item);
          if (replyIds.length > 0) {
            for (const rid of replyIds) {
              newItems.push({ type: 'commentRangeStart', id: rid });
              inserted = true;
            }
            newItems.push(item);
            // Adjacent del+ins replacement pair share author+date —
            // include the second half inside the comment range so we
            // don't break del-ins adjacency in the saved doc.
            const next = items[i + 1];
            if (
              next &&
              (item.type === 'insertion' || item.type === 'deletion') &&
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
  return inserted;
}

/**
 * Selection → live-region announcement derivation for the canvas renderer's
 * accessibility layer. Framework-free: the React `CanvasA11yLiveRegion` and
 * the Vue `CanvasA11yLiveRegion.vue` both derive their announcements here so
 * assistive-tech behavior cannot drift between adapters.
 *
 * The hidden textarea stays the focused editing surface in canvas mode, so
 * character/word echo while the caret moves comes from native input. The live
 * region only announces state *transitions* the visual layer would otherwise
 * communicate silently: a selection appearing or collapsing, the caret
 * entering commented / tracked-change / list content. Edge-triggered on the
 * selection context, so plain typing and uniform caret movement stay silent.
 *
 * @experimental part of the rust-canvas-engine change; shape may evolve.
 */

export interface A11ySelectionContext {
  hasSelection: boolean;
  inInsertion: boolean;
  inDeletion: boolean;
  inList: boolean;
  listLevel?: number;
  activeCommentIds: number[];
}

/** The slice of selection context the announcer reacts to. */
export interface A11ySelectionSnapshot {
  hasSelection: boolean;
  inInsertion: boolean;
  inDeletion: boolean;
  hasComment: boolean;
  inList: boolean;
  listLevel: number;
  /** display name of the active comment thread's author, when the host knows it */
  commentAuthor?: string;
  /** reply count of the active comment thread, when the host knows it */
  commentReplyCount?: number;
}

/**
 * host-supplied details of the comment thread(s) under the caret, resolved
 * from the same parsed comment model the sidebar uses. All strings are
 * file-derived — the live region renders them as text, never markup.
 */
export interface A11yCommentThreadDetails {
  authorName?: string;
  replyCount?: number;
}

export function snapshotFromSelectionContext(
  ctx: A11ySelectionContext,
  /** optional lookup so the `commentedText` announcement can carry author/reply vars */
  lookupCommentThread?: (commentIds: number[]) => A11yCommentThreadDetails | undefined
): A11ySelectionSnapshot {
  const thread =
    ctx.activeCommentIds.length > 0 ? lookupCommentThread?.(ctx.activeCommentIds) : undefined;
  return {
    hasSelection: ctx.hasSelection,
    inInsertion: ctx.inInsertion,
    inDeletion: ctx.inDeletion,
    hasComment: ctx.activeCommentIds.length > 0,
    inList: ctx.inList,
    listLevel: ctx.listLevel ?? 0,
    ...(thread?.authorName !== undefined ? { commentAuthor: thread.authorName } : {}),
    ...(thread?.replyCount !== undefined ? { commentReplyCount: thread.replyCount } : {}),
  };
}

/** i18n key under `a11y.*` plus interpolation vars */
export interface A11yAnnouncement {
  key:
    | 'selected'
    | 'selectionCleared'
    | 'commentedText'
    | 'trackedInsertion'
    | 'trackedDeletion'
    | 'listItem';
  vars?: Record<string, string | number>;
}

/**
 * announcements for one selection-state transition, in speaking order.
 * `prev === null` (first observation) announces nothing — there is no
 * transition to describe.
 */
export function computeA11yAnnouncements(
  prev: A11ySelectionSnapshot | null,
  next: A11ySelectionSnapshot
): A11yAnnouncement[] {
  if (!prev) return [];
  const out: A11yAnnouncement[] = [];
  if (!prev.hasSelection && next.hasSelection) out.push({ key: 'selected' });
  if (prev.hasSelection && !next.hasSelection) out.push({ key: 'selectionCleared' });
  if (!prev.hasComment && next.hasComment) {
    // thread details ride as interpolation vars when the host supplied them;
    // locales without {author}/{replies} placeholders simply ignore the vars
    out.push({
      key: 'commentedText',
      ...(next.commentAuthor !== undefined || next.commentReplyCount !== undefined
        ? {
            vars: {
              ...(next.commentAuthor !== undefined ? { author: next.commentAuthor } : {}),
              ...(next.commentReplyCount !== undefined ? { replies: next.commentReplyCount } : {}),
            },
          }
        : {}),
    });
  }
  if (!prev.inInsertion && next.inInsertion) out.push({ key: 'trackedInsertion' });
  if (!prev.inDeletion && next.inDeletion) out.push({ key: 'trackedDeletion' });
  if ((!prev.inList && next.inList) || (next.inList && prev.listLevel !== next.listLevel)) {
    out.push({ key: 'listItem', vars: { level: next.listLevel + 1 } });
  }
  return out;
}

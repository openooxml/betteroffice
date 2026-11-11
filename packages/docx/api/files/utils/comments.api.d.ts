// @betteroffice/docx declaration utils/comments.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Framework-agnostic comment + tracked-change helpers shared by the
 * React and Vue adapters. The data shapes and string-formatting rules
 * here are part of the visible UI (avatar colors, date strings) so
 * keep this file as the single source of truth — both adapters import
 * from here. CSS-property factories live in adapter-specific files
 * (sidebar/cardUtils.ts in React, sidebar/sidebarUtils.ts in Vue).
 * @packageDocumentation
 * @public
 */
import type { BlockContent, Comment, Paragraph } from '../types/content';
/** Extract plain text from a Comment's paragraph content. */
export declare function getCommentText(paragraphs?: Paragraph[] | BlockContent[]): string;
/** Extract text from the richest available comment body projection. */
export declare function getRichCommentText(comment: Pick<Comment, 'content' | 'blockContent'>): string;
export declare function formatDate(dateStr?: string): string;
export declare function getInitials(name: string): string;
export declare function getAvatarColor(name: string): string;
/** Prefer an authored reviewer color/palette assignment over name hashing. */
export declare function getReviewColor(reviewer: {
    name?: string;
    color?: string;
    paletteIndex?: number;
} | undefined): string;
/** Stable id shared by mirror anchors and framework comment cards. */
export declare function getCommentSemanticId(commentId: string | number): string;
export declare function truncateText(text: string, maxLength?: number): string;
/**
 * One tracked change surfaced by `extractTrackedChanges`. Each entry
 * groups all sites of one revision into a single row that the sidebar
 * renders as one card. Resolve via {@link acceptChangeById} /
 * {@link rejectChangeById} for any type — the by-id resolver walks
 * every site sharing the id so coalesced edits clear in one click.
 *
 * @public
 */
export interface TrackedChangeEntry {
    /**
     * Revision shape. Inline shapes (`insertion`, `deletion`, `replacement`)
     * wrap text runs; the rest are structural revisions on node attrs.
     *
     * - `insertion` — text was added (`<w:ins>`).
     * - `deletion` — text was struck through but not removed (`<w:del>`).
     * - `replacement` — a deletion + insertion by the same author at the
     *   same position+time; sidebar shows one combined card. `deletedText`
     *   and `insertionRevisionId` are set on this variant.
     * - `paragraphMarkInsertion` / `paragraphMarkDeletion` — Enter /
     *   Backspace produced a tracked paragraph break (`<w:pPr><w:rPr><w:ins/>` /
     *   `<w:del/>`).
     * - `paragraphPropertiesChanged` — formatting (alignment, spacing,
     *   etc.) on the paragraph was changed (`<w:pPrChange>`).
     * - `rowInserted` / `rowDeleted` / `rowPropertiesChanged` — table
     *   row authored / removed / formatted (`<w:trPr><w:ins/>` / `<w:del/>`
     *   / `<w:trPrChange>`).
     * - `cellInserted` / `cellDeleted` / `cellMerged` /
     *   `cellPropertiesChanged` — per-cell revisions
     *   (`<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>` / `<w:tcPrChange>`).
     * - `tablePropertiesChanged` — table-level formatting
     *   (`<w:tblPrChange>`).
     */
    type: 'insertion' | 'deletion' | 'replacement' | 'paragraphMarkInsertion' | 'paragraphMarkDeletion' | 'paragraphPropertiesChanged' | 'runPropertiesChanged' | 'rowInserted' | 'rowDeleted' | 'rowPropertiesChanged' | 'cellInserted' | 'cellDeleted' | 'cellMerged' | 'cellPropertiesChanged' | 'tableInserted' | 'tableDeleted' | 'tablePropertiesChanged';
    /**
     * Affected text. For inline types this is the run's text; for
     * structural types it's the surrounding paragraph / cell content
     * (truncated by the sidebar before display).
     */
    text: string;
    /**
     * Only set when `type === 'replacement'` — the text the user removed.
     * The inserted text lives in {@link TrackedChangeEntry.text}.
     */
    deletedText?: string;
    /** Author that minted the revision (`w:author`). */
    author: string;
    /** ISO timestamp the revision was minted (`w:date`). May be undefined for legacy imports. */
    date?: string;
    /**
     * Document position where the revision starts. For inline types this
     * is the start of the marked text run; for structural types it's the
     * containing paragraph / row / cell / table node's start position.
     * Used by the sidebar to anchor the card at the correct vertical
     * offset.
     */
    from: number;
    /**
     * Document position where the revision ends. For inline coalesced
     * runs that span multiple paragraphs, this is the END position of the
     * LAST run in the group; the intervening structural positions are not
     * preserved.
     */
    to: number;
    /**
     * The `w:id` of the revision. Pass to
     * {@link acceptChangeById} / {@link rejectChangeById} to resolve every
     * site sharing this id — including pPrIns paragraph attrs and
     * subsequent typed runs in the same editing session.
     */
    revisionId: number;
    /**
     * Only set when `type === 'replacement'` — the insertion half carries
     * a DIFFERENT `w:id` from the deletion (sharing would trip the OOXML
     * move-pair serializer). Card Accept handlers dispatch BOTH ids to
     * clear the deletion and the insertion + any coalesced paragraph-marks.
     */
    insertionRevisionId?: number;
    /**
     * Extra `w:id`s that map to the same logical revision as this card.
     * Populated when the extractor coalesces a burst of distinct ids by
     * (author, date) — e.g. a foreign document where the source editor
     * minted a fresh id per atomic edit. Accept/reject handlers must
     * resolve every id in this list in addition to {@link revisionId}.
     */
    coalescedRevisionIds?: number[];
}

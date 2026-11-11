// @betteroffice/docx declaration yrs/yrsToDocument.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Document, BlockContent, Paragraph, RunContent, TextFormatting } from '../types/document';
import type { YrsSession } from './index';
type Attrs = Record<string, unknown>;
/** @internal */
export declare function stableStringify(value: unknown): string;
/** @internal */
export declare function stableStringifyWithout(value: Attrs, omitted: ReadonlySet<string>): string;
/** @internal */
export declare function runContentForText(text: string, formatting: TextFormatting | (() => TextFormatting)): RunContent[];
/** Serialized ids for a session revision's deletion and insertion. */
export interface RevisionNumbers {
    insertion?: number;
    deletion?: number;
}
export declare function assertCommentBookmarkReplacement(blocks: readonly BlockContent[]): void;
export declare function restoreBlockBookmarks(projected: BlockContent[], base: readonly BlockContent[], storyId: string): BlockContent[];
/** The session key of the pilcrow a projected paragraph was built from. @internal */
export declare function projectedSessionKey(paragraph: Paragraph): string | undefined;
/**
 * `blocks` with paragraphs of their own, in table cells too, each still the
 * session paragraph it was projected from: a save edits these, not the
 * projection the next one reuses. @internal
 */
export declare function ownProjectedParagraphs(blocks: readonly BlockContent[]): BlockContent[];
/** Stories a committed op touched; `all` invalidates the whole session. */
export declare function noteYrsStoriesDirty(session: YrsSession, stories: 'all' | string | Iterable<string>): void;
/**
 * Rebuild all yrs-owned editable stories into the serializer-facing Document
 * while preserving every package part the editor does not own.
 */
export interface YrsToDocumentOptions {
    /**
     * Root stories to project. Omit to rebuild every editable story. A `body`
     * projection includes its nested table/content-control stories.
     */
    storyIds?: ReadonlySet<string>;
    /**
     * The session comment id of each comment in `base` whose OOXML id is not
     * that id, such as one added with {@link YrsSession.addComment}. @internal
     */
    commentIds?: ReadonlyMap<number, string>;
}
export declare function yrsToDocument(session: YrsSession, base: Document, options?: YrsToDocumentOptions): Document;
/**
 * {@link yrsToDocument} of the body alone, writing each session revision with the serialized
 * ids `revisionIds` gives its deletion and insertion.
 */
export declare function yrsBodyToDocumentWithRevisionIds(session: YrsSession, base: Document, revisionIds: ReadonlyMap<string, RevisionNumbers>): Document;
export {};

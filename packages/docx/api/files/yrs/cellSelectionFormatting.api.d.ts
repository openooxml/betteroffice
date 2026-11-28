// @betteroffice/docx declaration yrs/cellSelectionFormatting.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsAuthor, YrsSession, YrsStoryRange, YrsTableRange } from './index';
/** A run mark the selection context reports as on, off or mixed. @internal */
export type YrsToggledMark = 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript';
/** Whether the caret belongs to a cell covered by a multi-cell selection. @internal */
export declare function cellSelectionContainsCaret(reader: Pick<YrsSession, 'tablePayload'>, cells: YrsTableRange, caretStory: string): boolean;
/** The whole text of each non-empty cell that `cells` covers, in reading order. @internal */
export declare function cellSelectionRanges(reader: Pick<YrsSession, 'tablePayload' | 'paragraphSpans'>, cells: YrsTableRange): YrsStoryRange[];
/**
 * Sets the same mark direction in every range, using the direct cell text to choose it.
 * @internal
 */
export declare function toggleAcross(session: Pick<YrsSession, 'paragraphs' | 'selectionContext' | 'formatRange'>, ranges: readonly YrsStoryRange[], mark: YrsToggledMark, suggesting?: YrsAuthor): void;

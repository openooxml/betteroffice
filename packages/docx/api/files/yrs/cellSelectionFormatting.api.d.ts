// @betteroffice/docx declaration yrs/cellSelectionFormatting.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsSession, YrsStoryRange, YrsTableRange } from './index';
/** A run mark the selection context reports as on, off or mixed. @internal */
export type YrsToggledMark = 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript';
/** The whole text of each non-empty cell that `cells` covers, in reading order. @internal */
export declare function cellSelectionRanges(reader: Pick<YrsSession, 'tablePayload' | 'paragraphSpans'>, cells: YrsTableRange): YrsStoryRange[];
/**
 * Toggles `mark` over `ranges` as one selection: off where it is set once all their text carries
 * it, otherwise on where it is not. A range whose story holds no text, such as an empty cell,
 * does not keep it from turning off.
 * @internal
 */
export declare function toggleAcross(session: Pick<YrsSession, 'paragraphs' | 'selectionContext'>, ranges: readonly YrsStoryRange[], mark: YrsToggledMark, toggle: (range: YrsStoryRange, on: boolean) => void): void;

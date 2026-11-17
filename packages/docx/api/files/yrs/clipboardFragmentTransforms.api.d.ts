// @betteroffice/docx declaration yrs/clipboardFragmentTransforms.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxClipboardFragment } from './clipboardFragment';
/** @internal The three variants a paste can be re-applied as. */
export type PasteFormattingOption = 'keep-source' | 'merge' | 'text-only';
/**
 * @internal Merge-formatting variant of a pasted fragment: bold, italic,
 * underline, strikethrough, hyperlinks and list numbering survive; font
 * family, size, colour, highlight and paragraph style are dropped so the
 * destination paragraph's formatting applies.
 */
export declare function mergeFormattingFragment(fragment: DocxClipboardFragment): DocxClipboardFragment;
/**
 * @internal Text-only variant of a pasted fragment: one plain paragraph per
 * source paragraph, with no run, paragraph or list formatting of its own.
 */
export declare function textOnlyFragment(fragment: DocxClipboardFragment): DocxClipboardFragment;
/** @internal Whether a fragment carries any formatting a paste variant could change. */
export declare function clipboardFragmentHasFormatting(fragment: DocxClipboardFragment | null): boolean;

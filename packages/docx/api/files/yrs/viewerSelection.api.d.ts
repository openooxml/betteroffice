// @betteroffice/docx declaration yrs/viewerSelection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxTextRange } from './edits';
import type { YrsLoc } from './index';
import { DisplayPositionIndex } from './displayPositionIndex';
/** @internal How far a selection grows around a display position. */
export type DocxSelectionUnit = 'word' | 'paragraph' | 'story';
/** @internal A selection in the display positions of one root story's layout. */
export interface DocxDisplayRange {
    anchor: number;
    head: number;
}
/** @internal The text a display selection covers. */
export interface DocxDisplaySelectionText {
    /** Plain text as the clipboard takes it: tabs and breaks as characters, tables as tab-separated rows. */
    text: string;
    /** The accepted-view range, or null when the selection spans stories. */
    range: DocxTextRange | null;
}
export interface DocxDisplaySelectionInfo {
    paraId: string | null;
    selectedText: string;
    paragraphText: string;
    before: string;
    after: string;
}
/** @internal The word, paragraph or whole story around `position`, in display positions. */
export declare function resolveSelectionUnit(index: DisplayPositionIndex, rootStory: string, position: number, unit: DocxSelectionUnit, expectVersion: string): DocxDisplayRange | null;
/** @internal The plain text and accepted-view range of the selection `[anchor, head]`. */
export declare function resolveSelectionText(index: DisplayPositionIndex, rootStory: string, anchor: number, head: number, expectVersion: string): DocxDisplaySelectionText | null;
/** @internal The selection info for a display range. */
export declare function resolveSelectionInfo(index: DisplayPositionIndex, rootStory: string, anchor: number, head: number, expectVersion: string): DocxDisplaySelectionInfo | null;
/** @internal The paragraph range across stories, mapped into the root layout. */
export declare function resolveParagraphTarget(index: DisplayPositionIndex, rootStory: string, paraId: string, expectVersion: string): DocxDisplayRange | null;
/** @internal The first comment anchor in the root story's display positions. */
export declare function resolveCommentTarget(index: DisplayPositionIndex, rootStory: string, commentId: string, expectVersion: string): DocxDisplayRange | null;
/** @internal The revision range in the root story's display positions. */
export declare function resolveRevisionTarget(index: DisplayPositionIndex, rootStory: string, revisionId: string, expectVersion: string): DocxDisplayRange | null;
/** @internal The display position of bookmark `name` in root story `rootStory`, or null. */
export declare function resolveBookmarkPosition(index: DisplayPositionIndex, rootStory: string, name: string, expectVersion: string): number | null;
export declare function resolveBookmarkTarget(index: DisplayPositionIndex, rootStory: string, name: string, expectVersion: string): {
    loc: YrsLoc;
    position: number;
} | null;

// @betteroffice/docx declaration yrs/clipboardFragment.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsAuthor, YrsLoc, YrsSession, YrsStoryRange, YrsTableRange, YrsInlineFormatDelta } from './index';
import { type TablePayload } from './storyPlainText';
/** @internal Version of the internal clipboard payload; a reader ignores other versions. */
export declare const DOCX_CLIPBOARD_FRAGMENT_VERSION = 2;
/**
 * @internal One unit of a table cell's story. `attributes` are raw yrs run or embed attributes as
 * `storySegments` reports them; `format` is a semantic patch applied on top (HTML import uses it).
 */
export type DocxFragmentSegment = {
    kind: 'text';
    text: string;
    attributes?: Record<string, unknown>;
    format?: YrsInlineFormatDelta;
} | {
    kind: 'tab' | 'break';
    payload?: Record<string, unknown>;
    attributes?: Record<string, unknown>;
} | {
    kind: 'pilcrow';
    properties?: Record<string, unknown>;
    attributes?: Record<string, unknown>;
} | {
    kind: 'table';
    table: DocxTableFragment;
    attributes?: Record<string, unknown>;
};
/** @internal A cell's `content` ends with a pilcrow; `shading` is a hex fill applied like `setCellShading`. */
export interface DocxTableFragmentCell {
    tcPr: Record<string, unknown>;
    content: DocxFragmentSegment[];
    shading?: string;
}
/** @internal */
export interface DocxTableFragmentRow {
    trPr: Record<string, unknown>;
    cells: DocxTableFragmentCell[];
}
/** @internal A rectangular table, its cells in row order with `rowspan`/`colspan` in `tcPr`. */
export interface DocxTableFragment {
    tblPr: Record<string, unknown>;
    /** Column widths in twips; empty splits the default table width evenly. */
    grid: number[];
    rows: DocxTableFragmentRow[];
}
export type DocxClipboardInline = {
    text: string;
} | {
    image: Record<string, unknown>;
};
/**
 * @internal One formatted run of a clipboard paragraph; the runs of a paragraph join to its `text`.
 * A run with `format.other` replaces the character formatting at the paste target, as Keep Source does.
 */
export interface DocxClipboardRun {
    text: string;
    format?: YrsInlineFormatDelta;
    href?: string;
    tooltip?: string;
}
/** @internal A paragraph's leading tables and ordered inline content. With `inlines`, `runs` pairs one to one with its text entries. */
export interface DocxClipboardParagraph {
    tables?: DocxTableFragment[];
    text: string;
    inlines?: DocxClipboardInline[];
    runs?: DocxClipboardRun[];
}
/** @internal Copied or pasted content in document order; consecutive paragraphs are separate paragraphs. */
export interface DocxClipboardFragment {
    paragraphs: DocxClipboardParagraph[];
}
/** @internal What a copy writes: plain text plus a fragment, null when the selection exceeds the fragment limits. */
export interface DocxClipboardSlice {
    text: string;
    fragment: DocxClipboardFragment | null;
}
/** @internal */
export declare function clipboardFragmentHasPictures(fragment: DocxClipboardFragment | null): boolean;
/** @internal Whether a fragment holds only text, so it can paste wherever text can. */
export declare function clipboardFragmentIsTextOnly(fragment: DocxClipboardFragment): boolean;
/** @internal A paragraph's inline content in order, its text as formatted runs where it has them. */
export declare function clipboardParagraphUnits(paragraph: DocxClipboardParagraph): Array<DocxClipboardRun | {
    image: Record<string, unknown>;
}>;
/** @internal Strict shape and size check for a pasted or copied clipboard fragment. */
export declare function validClipboardFragment(value: unknown): value is DocxClipboardFragment;
/** @internal Sessions a table read needs: the raw payload plus story segments. */
export type ClipboardTableReader = Pick<YrsSession, 'tablePayload' | 'storySegments'> & {
    mediaDataUrl?(source: string): string | null | undefined;
};
/** @internal Engine surface a fragment paste needs. */
export type ClipboardPasteEngine = Pick<YrsSession, 'deleteRange' | 'splitParagraph' | 'insertText' | 'insertTableFragment' | 'validateTableFragment' | 'insertImage' | 'clearFormatting' | 'formatRange' | 'setHyperlink'>;
/** @internal Inclusive grid bounds for {@link readTableFragment}. */
export interface ClipboardTableRect {
    top: number;
    bottom: number;
    left: number;
    right: number;
}
/** @internal Copy source: a multi-cell range, or story units `[from, to)`. */
export type ClipboardSliceTarget = {
    cells: YrsTableRange;
} | {
    story: string;
    from: number;
    to: number;
};
/** @internal Raw table payload lowered to the anchor reader's shape, or null. */
export declare function clipboardTablePayload(value: unknown): TablePayload | null;
/**
 * Reads one rectangular table as a fragment. `rect` is inclusive; it expands
 * to fully cover every merged cell it touches.
 * @internal
 */
export declare function readTableFragment(reader: ClipboardTableReader, table: {
    story: string;
    tableIndex: number;
}, rect?: ClipboardTableRect): DocxTableFragment;
/**
 * Reads the clipboard slice for a copy: always plain text, plus a fragment
 * when the selection holds tables or pictures.
 * @internal
 */
export declare function readClipboardSlice(reader: ClipboardTableReader, target: ClipboardSliceTarget, themeColors?: Readonly<Record<string, string>>): DocxClipboardSlice;
/**
 * Pastes a fragment over a story range, returning the caret. Deletes a
 * non-collapsed range first, then replays each paragraph: splits for
 * paragraphs after the first, inserts tables, then text.
 * @internal
 */
export declare function applyClipboardFragment(engine: ClipboardPasteEngine, range: YrsStoryRange, fragment: DocxClipboardFragment, suggesting?: YrsAuthor): YrsLoc;

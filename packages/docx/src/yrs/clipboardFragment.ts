import type { YrsInlineFormatDelta } from './index';

/** @internal Version of the internal clipboard payload; a reader ignores other versions. */
export const DOCX_CLIPBOARD_FRAGMENT_VERSION = 1;

/**
 * @internal One unit of a table cell's story. `attributes` are raw yrs run or embed attributes as
 * `storySegments` reports them; `format` is a semantic patch applied on top (HTML import uses it).
 */
export type DocxFragmentSegment =
  | { kind: 'text'; text: string; attributes?: Record<string, unknown>; format?: YrsInlineFormatDelta }
  | { kind: 'tab' | 'break'; payload?: Record<string, unknown>; attributes?: Record<string, unknown> }
  | { kind: 'pilcrow'; properties?: Record<string, unknown>; attributes?: Record<string, unknown> }
  | { kind: 'table'; table: DocxTableFragment; attributes?: Record<string, unknown> };

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

/** @internal A paragraph's leading tables, then its text (tabs allowed, no line breaks). */
export interface DocxClipboardParagraph {
  tables?: DocxTableFragment[];
  text: string;
}

/** @internal Copied or pasted content in document order; consecutive paragraphs are separate paragraphs. */
export interface DocxClipboardFragment {
  paragraphs: DocxClipboardParagraph[];
}

/** @internal What a copy writes: always plain text, and a fragment when the selection holds a table. */
export interface DocxClipboardSlice {
  text: string;
  fragment: DocxClipboardFragment | null;
}

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

const MAX_FRAGMENT_CELLS = 10_000;
const MAX_FRAGMENT_PARAGRAPHS = 1_000;
const MAX_FRAGMENT_TEXT = 1_000_000;
const MAX_FRAGMENT_DEPTH = 8;
const LINE_BREAK = /[\n\r\u000B  ]/u;

function plainObject(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function jsonValue(value: unknown, depth = 0): boolean {
  if (depth > 32) return false;
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return true;
  if (typeof value === 'number') return Number.isFinite(value);
  if (Array.isArray(value)) return value.every(entry => jsonValue(entry, depth + 1));
  return plainObject(value) && Object.values(value).every(entry => jsonValue(entry, depth + 1));
}

function optionalMap(value: unknown): boolean {
  return value === undefined || (plainObject(value) && jsonValue(value));
}

/** @internal Strict shape and size check for a pasted or copied clipboard fragment. */
export function validClipboardFragment(value: unknown): value is DocxClipboardFragment {
  const budget = { cells: MAX_FRAGMENT_CELLS, text: MAX_FRAGMENT_TEXT };
  const text = (entry: unknown, allowEmpty: boolean): boolean => {
    if (typeof entry !== 'string' || (!allowEmpty && entry === '') || LINE_BREAK.test(entry)) return false;
    budget.text -= entry.length;
    return budget.text >= 0;
  };
  const segment = (entry: unknown, depth: number): boolean => {
    if (!plainObject(entry) || !optionalMap(entry.attributes)) return false;
    switch (entry.kind) {
      case 'text': return text(entry.text, false) && optionalMap(entry.format);
      case 'tab': case 'break': return optionalMap(entry.payload);
      case 'pilcrow': return optionalMap(entry.properties);
      case 'table': return table(entry.table, depth + 1);
      default: return false;
    }
  };
  const table = (entry: unknown, depth: number): boolean => {
    if (depth > MAX_FRAGMENT_DEPTH || !plainObject(entry) || !optionalMap(entry.tblPr) || entry.tblPr === undefined) return false;
    if (!Array.isArray(entry.grid) || !entry.grid.every(width => typeof width === 'number' && Number.isFinite(width) && width > 0)) return false;
    if (!Array.isArray(entry.rows) || entry.rows.length === 0 || entry.rows.length > 1000) return false;
    return entry.rows.every(row => plainObject(row) && row.trPr !== undefined && optionalMap(row.trPr) &&
      Array.isArray(row.cells) && row.cells.length > 0 && row.cells.length <= 1000 && row.cells.every(cell => {
        if (!plainObject(cell) || cell.tcPr === undefined || !optionalMap(cell.tcPr)) return false;
        if (cell.shading !== undefined && (typeof cell.shading !== 'string' || !/^[0-9a-f]{6}$/iu.test(cell.shading))) return false;
        if (--budget.cells < 0 || !Array.isArray(cell.content) || cell.content.length === 0) return false;
        const last = cell.content.at(-1);
        return plainObject(last) && last.kind === 'pilcrow' && cell.content.every(item => segment(item, depth));
      }));
  };
  if (!plainObject(value) || !Array.isArray(value.paragraphs) || value.paragraphs.length === 0 ||
    value.paragraphs.length > MAX_FRAGMENT_PARAGRAPHS) return false;
  return value.paragraphs.every(paragraph => plainObject(paragraph) && text(paragraph.text, true) &&
    (paragraph.tables === undefined || (Array.isArray(paragraph.tables) && paragraph.tables.every(entry => table(entry, 1)))));
}

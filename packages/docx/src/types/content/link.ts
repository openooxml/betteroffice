/**
 * Hyperlinks (`w:hyperlink`), bookmark markers (`w:bookmarkStart`/`End`),
 * and field types (`w:fldSimple`, complex `w:fldChar` runs).
 */

import type { Run } from './run';
import type { TextFormatting } from '../formatting';
import type { InlineSdt } from './sdt';
import type { MathEquation } from './math';
import type { BlockContent } from './section';

/** Stable document position used by bookmark and field range contracts. */
export interface ContentPosition {
  /** Stable block/paragraph id. Undefined = unresolved. */
  blockId?: string | number;
  /** Inline offset within the block. Undefined = block boundary. */
  offset?: number;
  /** Table/grid ancestry for cross-cell ranges. Undefined = body flow. */
  path?: number[];
}

/** A resolved bookmark range; absent endpoints represent an unresolved marker. */
export interface BookmarkRange {
  id?: number;
  name?: string;
  start?: ContentPosition;
  end?: ContentPosition;
  colFirst?: number;
  colLast?: number;
}

/**
 * Hyperlink (`w:hyperlink`) — wraps runs in a clickable link. External
 * targets resolve through the relationships part (`rId` → `href`);
 * internal targets reference a `BookmarkStart` anchor by name.
 */
export interface Hyperlink {
  type: 'hyperlink';
  /** Relationship ID for external link */
  rId?: string;
  /** Resolved URL (from relationships) */
  href?: string;
  /** Internal bookmark anchor */
  anchor?: string;
  /** Tooltip text */
  tooltip?: string;
  /** Target frame */
  target?: string;
  /** Link history tracking */
  history?: boolean;
  /** Document location */
  docLocation?: string;
  /** Child runs */
  children: (Run | BookmarkStart | BookmarkEnd)[];
  /**
   * Recursive `EG_PContent` projection for fields, SDTs, math, and bookmarks.
   * Undefined means consumers use the legacy `children` run list.
   */
  structuredChildren?: HyperlinkContent[];
}

/** Recursive hyperlink child grammar. Nested hyperlinks remain invalid. */
export type HyperlinkContent =
  | Run
  | BookmarkStart
  | BookmarkEnd
  | SimpleField
  | ComplexField
  | InlineSdt
  | MathEquation;

/**
 * Bookmark start marker (w:bookmarkStart)
 */
export interface BookmarkStart {
  type: 'bookmarkStart';
  /** Bookmark ID */
  id: number;
  /** Bookmark name */
  name: string;
  /** Column index for table bookmarks */
  colFirst?: number;
  colLast?: number;
  /** Exact source position. Undefined = legacy paragraph-boundary behavior. */
  position?: ContentPosition;
}

/**
 * Bookmark end marker (w:bookmarkEnd)
 */
export interface BookmarkEnd {
  type: 'bookmarkEnd';
  /** Bookmark ID */
  id: number;
  /** Exact source position. Undefined = legacy paragraph-boundary behavior. */
  position?: ContentPosition;
}

/**
 * Known field types
 */
export type FieldType =
  | 'PAGE'
  | 'NUMPAGES'
  | 'NUMWORDS'
  | 'NUMCHARS'
  | 'DATE'
  | 'TIME'
  | 'CREATEDATE'
  | 'SAVEDATE'
  | 'PRINTDATE'
  | 'AUTHOR'
  | 'TITLE'
  | 'SUBJECT'
  | 'KEYWORDS'
  | 'COMMENTS'
  | 'FILENAME'
  | 'FILESIZE'
  | 'TEMPLATE'
  | 'DOCPROPERTY'
  | 'DOCVARIABLE'
  | 'REF'
  | 'PAGEREF'
  | 'NOTEREF'
  | 'HYPERLINK'
  | 'TOC'
  | 'TOA'
  | 'INDEX'
  | 'SEQ'
  | 'STYLEREF'
  | 'AUTONUM'
  | 'AUTONUMLGL'
  | 'AUTONUMOUT'
  | 'IF'
  | 'MERGEFIELD'
  | 'NEXT'
  | 'NEXTIF'
  | 'ASK'
  | 'SET'
  | 'QUOTE'
  | 'INCLUDETEXT'
  | 'INCLUDEPICTURE'
  | 'SYMBOL'
  | 'ADVANCE'
  | 'EDITTIME'
  | 'REVNUM'
  | 'SECTION'
  | 'SECTIONPAGES'
  | 'USERADDRESS'
  | 'USERNAME'
  | 'USERINITIALS'
  | 'FORMTEXT'
  | 'FORMCHECKBOX'
  | 'FORMDROPDOWN'
  | 'CITATION'
  | 'BIBLIOGRAPHY'
  | 'UNKNOWN';

/**
 * Recursive, inert field content. Undefined members are empty; no instruction
 * in this contract is executable merely because it was parsed.
 */
export interface StructuredFieldContent {
  /** Inline result/code children in source order. */
  inline?: FieldInlineContent[];
  /** Multi-paragraph/table result children in source order. */
  blocks?: BlockContent[];
}

/** Field-capable inline grammar used by nested field code/results. */
export type FieldInlineContent =
  | Run
  | Hyperlink
  | InlineSdt
  | MathEquation
  | SimpleField
  | ComplexField;

/** Versioned structured field node; missing version means legacy v0 fields. */
export interface StructuredFieldTree {
  /** Contract version. Undefined reads as 0. */
  version?: number;
  /** Parsed code container. Undefined = use legacy code fields. */
  code?: StructuredFieldContent;
  /** Cached result container. Undefined = use legacy result fields. */
  result?: StructuredFieldContent;
  /** Nested child fields in document order. */
  children?: StructuredFieldTree[];
  /** Result/code presentation. Undefined = result. */
  displayMode?: 'result' | 'code';
}

/**
 * Simple field (w:fldSimple)
 */
export interface SimpleField {
  type: 'simpleField';
  /** Field instruction (e.g., "PAGE \\* MERGEFORMAT") */
  instruction: string;
  /** Parsed field type */
  fieldType: FieldType;
  /** Current display value */
  content: (Run | Hyperlink)[];
  /** Field is locked */
  fldLock?: boolean;
  /** Field is dirty */
  dirty?: boolean;
  /** Rich recursive result projection. Undefined = use `content`. */
  structuredResult?: StructuredFieldContent;
  /** Parsed/versioned field tree. Undefined = legacy flat field. */
  fieldTree?: StructuredFieldTree;
}

/**
 * Complex field (w:fldChar begin/separate/end with w:instrText)
 */
export interface ComplexField {
  type: 'complexField';
  /** Field instruction */
  instruction: string;
  /** Parsed field type */
  fieldType: FieldType;
  /** Field code runs */
  fieldCode: Run[];
  /** Display result runs */
  fieldResult: Run[];
  /**
   * Run formatting carried by the field's structural runs (the runs holding
   * the `w:fldChar` begin/separate/end). Word styles the field result with
   * this `w:rPr` when there is no separate result run (e.g. a `PAGE` field
   * collapsed into a single run). Used as a fallback for rendering and
   * serialization so the formatting survives the round-trip.
   */
  formatting?: TextFormatting;
  /** Field is locked */
  fldLock?: boolean;
  /** Field is dirty */
  dirty?: boolean;
  /** Rich recursive code projection. Undefined = use `fieldCode`. */
  structuredCode?: StructuredFieldContent;
  /** Rich recursive cached result. Undefined = use `fieldResult`. */
  structuredResult?: StructuredFieldContent;
  /** Parsed/versioned nested field tree. Undefined = legacy flat field. */
  fieldTree?: StructuredFieldTree;
}

export type Field = SimpleField | ComplexField;

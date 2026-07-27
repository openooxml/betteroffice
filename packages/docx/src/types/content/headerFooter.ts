/**
 * Page furniture — headers (`w:hdr`), footers (`w:ftr`), footnotes
 * (`w:footnote`), and endnotes (`w:endnote`), plus the section-level
 * properties (`w:footnotePr`/`w:endnotePr`) that configure note layout.
 */

import type { NumberFormat } from '../lists';
import type { BlockContent } from './section';
import type { Watermark } from './watermark';

/**
 * Header/footer type
 */
export type HeaderFooterType = 'default' | 'first' | 'even';

/**
 * Header or footer reference
 */
export interface HeaderReference {
  type: HeaderFooterType;
  rId: string;
}

export interface FooterReference {
  type: HeaderFooterType;
  rId: string;
}

/**
 * Header or footer content
 */
export interface HeaderFooter {
  type: 'header' | 'footer';
  /** Header/footer type */
  hdrFtrType: HeaderFooterType;
  /** Content (paragraphs, tables, etc.) */
  content: BlockContent[];
  /**
   * Watermark stored on this header (MS Word "Design → Watermark"). Lives
   * here, not in `content`, so it stays out of the editable text flow while
   * still round-tripping. Only headers carry watermarks; footers never do.
   */
  watermark?: Watermark;
}

/**
 * Footnote position
 */
export type FootnotePosition = 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd';

/**
 * Endnote position
 */
export type EndnotePosition = 'sectEnd' | 'docEnd';

/**
 * Number restart type
 */
export type NoteNumberRestart = 'continuous' | 'eachSect' | 'eachPage';

/** Shared note kind used by footnote/endnote layout contracts. */
export type NoteKind = 'footnote' | 'endnote';

/** Reference to a special separator/continuation note. */
export interface NoteSeparatorReference {
  kind?: 'separator' | 'continuationSeparator' | 'continuationNotice';
  noteId?: number;
}

/**
 * Footnote properties
 */
export interface FootnoteProperties {
  position?: FootnotePosition;
  numFmt?: NumberFormat;
  numStart?: number;
  numRestart?: NoteNumberRestart;
  /** Optional custom number-format string (`w:numFmt/@w:format`). */
  customNumberFormat?: string;
  /** Authored special-note references. Undefined = package defaults. */
  separators?: NoteSeparatorReference[];
}

/**
 * Endnote properties
 */
export interface EndnoteProperties {
  position?: EndnotePosition;
  numFmt?: NumberFormat;
  numStart?: number;
  numRestart?: NoteNumberRestart;
  /** Optional custom number-format string (`w:numFmt/@w:format`). */
  customNumberFormat?: string;
  /** Authored special-note references. Undefined = package defaults. */
  separators?: NoteSeparatorReference[];
}

/**
 * Footnote (w:footnote)
 */
export interface Footnote {
  type: 'footnote';
  /** Footnote ID */
  id: number;
  /** Special footnote type */
  noteType?: 'normal' | 'separator' | 'continuationSeparator' | 'continuationNotice';
  /**
   * Content. Per ECMA-376 §17.11.10 footnotes can hold the same blocks as
   * the body, so the note parser reuses the body's `parseBlockContent`: the
   * full block model — paragraphs, tables, and block-level `w:sdt` content
   * controls (as `BlockSdt`) — flows through the resident Rust layout path
   * and stays editable on round-trip.
   */
  content: BlockContent[];
  /** Verbatim note XML for unmodeled bookmarks or custom XML. */
  verbatimXml?: string;
}

/**
 * Endnote (w:endnote)
 */
export interface Endnote {
  type: 'endnote';
  /** Endnote ID */
  id: number;
  /** Special endnote type */
  noteType?: 'normal' | 'separator' | 'continuationSeparator' | 'continuationNotice';
  /**
   * Content. Per ECMA-376 §17.11.4 endnotes can hold the same blocks as
   * the body — paragraphs, tables, and block-level content controls. See note
   * on `Footnote.content`.
   */
  content: BlockContent[];
  /** Verbatim note XML for unmodeled blocks. */
  verbatimXml?: string;
}

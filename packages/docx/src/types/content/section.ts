/**
 * Section properties (`w:sectPr`) — page size and margins, columns,
 * header/footer references, line numbers, page borders, document grid,
 * paper sources — plus the section and document-body containers that
 * group block-level content.
 */

import type { ColorValue, ThemeColorSlot, BorderSpec } from '../colors';
import type { Paragraph } from './paragraph';
import type { Table } from './table';
import type { BlockSdt } from './sdt';
import type {
  HeaderFooter,
  HeaderFooterType,
  HeaderReference,
  FooterReference,
  FootnoteProperties,
  EndnoteProperties,
} from './headerFooter';
import type { Comment } from './comment';
import type { CommentAuthor } from './comment';
import type { NumberFormat } from '../lists';
import type { Watermark } from './watermark';

/**
 * Page orientation
 */
export type PageOrientation = 'portrait' | 'landscape';

/**
 * Section start type
 */
export type SectionStart = 'continuous' | 'nextPage' | 'oddPage' | 'evenPage' | 'nextColumn';

/**
 * Vertical alignment
 */
export type VerticalAlign = 'top' | 'center' | 'both' | 'bottom';

/**
 * Line number restart type
 */
export type LineNumberRestart = 'continuous' | 'newPage' | 'newSection';

/**
 * Column definition
 */
export interface Column {
  /** Column width in twips */
  width?: number;
  /** Space after column in twips */
  space?: number;
}

/** Section-scoped page label settings from `w:pgNumType`. */
export interface PageNumberingProperties {
  /** Logical page-number restart (`@w:start`). Undefined = continue. */
  start?: number;
  /** Number format (`@w:fmt`). Undefined = decimal. */
  format?: NumberFormat;
  /** Heading/list outline level used for chapter prefixes (`@w:chapStyle`). */
  chapterStyle?: number;
  /** Chapter/page separator (`@w:chapSep`). Undefined = hyphen. */
  chapterSeparator?: 'colon' | 'emDash' | 'enDash' | 'hyphen' | 'period';
}

/**
 * Section properties (`w:sectPr`) — page geometry, margins, columns,
 * header/footer references, and page numbering for one section of the
 * document. Sections are introduced by inline `sectPr` markers on the
 * terminating paragraph (`Paragraph.sectionProperties`) and the body's
 * final `sectPr`.
 *
 * All distance units are twips (1/20 of a point) on the wire. The layout
 * engine converts to pixels.
 *
 * See ECMA-376 §17.6.
 */
export interface SectionProperties {
  /** Stable section identity. Undefined preserves positional identity. */
  sectionId?: string;
  // Page size
  /** Page width in twips */
  pageWidth?: number;
  /** Page height in twips */
  pageHeight?: number;
  /** Page orientation */
  orientation?: PageOrientation;

  // Margins
  /** Top margin in twips */
  marginTop?: number;
  /** Bottom margin in twips */
  marginBottom?: number;
  /** Left margin in twips */
  marginLeft?: number;
  /** Right margin in twips */
  marginRight?: number;
  /** Header distance from top in twips */
  headerDistance?: number;
  /** Footer distance from bottom in twips */
  footerDistance?: number;
  /** Gutter margin in twips */
  gutter?: number;

  // Columns
  /** Number of columns */
  columnCount?: number;
  /** Space between columns in twips */
  columnSpace?: number;
  /** Equal width columns */
  equalWidth?: boolean;
  /** Separator line between columns */
  separator?: boolean;
  /** Individual column definitions */
  columns?: Column[];
  /**
   * Number of columns the footnote area is laid out in (`w15:footnoteColumns`).
   * Word's "Footnote layout → Columns" setting, independent of the body column
   * count above. Undefined/1 means the footnote area follows the body (single
   * column for a single-column section). See ECMA-376 + the w15 extension.
   */
  footnoteColumns?: number;

  // Section behavior
  /** Section start type */
  sectionStart?: SectionStart;
  /** Vertical alignment of text */
  verticalAlign?: VerticalAlign;
  /** Right-to-left section */
  bidi?: boolean;

  // Headers and footers
  /** Header references */
  headerReferences?: HeaderReference[];
  /** Footer references */
  footerReferences?: FooterReference[];
  /** Different first page header/footer */
  titlePg?: boolean;
  /** Different odd/even page headers/footers */
  evenAndOddHeaders?: boolean;

  // Line numbers
  /** Line numbering settings */
  lineNumbers?: {
    start?: number;
    countBy?: number;
    distance?: number;
    restart?: LineNumberRestart;
  };

  /** Section-scoped page-number restart/format (`w:pgNumType`). */
  pageNumbering?: PageNumberingProperties;

  // Page borders
  /** Page borders */
  pageBorders?: {
    top?: BorderSpec;
    bottom?: BorderSpec;
    left?: BorderSpec;
    right?: BorderSpec;
    /** Display setting */
    display?: 'allPages' | 'firstPage' | 'notFirstPage';
    /** Offset from */
    offsetFrom?: 'page' | 'text';
    /** Z-order */
    zOrder?: 'front' | 'back';
  };

  // Background
  /** Page background */
  background?: {
    color?: ColorValue;
    themeColor?: ThemeColorSlot;
    themeTint?: string;
    themeShade?: string;
  };

  /** Effective header-owned watermark projection. Undefined = none/inherit. */
  watermark?: Watermark;

  // Footnote/Endnote properties
  /** Footnote properties for this section */
  footnotePr?: FootnoteProperties;
  /** Endnote properties for this section */
  endnotePr?: EndnoteProperties;

  // Document grid
  /** Document grid */
  docGrid?: {
    type?: 'default' | 'lines' | 'linesAndChars' | 'snapToChars';
    linePitch?: number;
    charSpace?: number;
  };

  // Paper source
  /** First page paper source */
  paperSrcFirst?: number;
  /** Other pages paper source */
  paperSrcOther?: number;
}

/**
 * Block-level content types
 */
export type BlockContent = Paragraph | Table | BlockSdt;

/**
 * One section of the document — a `SectionProperties` plus the block
 * content (`Paragraph`s and `Table`s) that lives under those properties.
 *
 * Sections are derived during parse: every paragraph carrying an inline
 * `sectPr` ends a section, and the body's final `sectPr` defines the
 * last section. Each section may carry its own headers/footers map.
 */
export interface Section {
  /** Stable section identity. Undefined = derive from array position. */
  id?: string;
  /** Section properties */
  properties: SectionProperties;
  /** Content in this section */
  content: BlockContent[];
  /** Headers for this section */
  headers?: Map<HeaderFooterType, HeaderFooter>;
  /** Footers for this section */
  footers?: Map<HeaderFooterType, HeaderFooter>;
}

/**
 * Document body (`w:body`) — the editable content of the document.
 *
 * Contains the ordered block content (paragraphs and tables), the section
 * layout chain derived from inline `sectPr` markers, the final `sectPr`,
 * and any document-level comments. This is what most edit operations
 * mutate; headers/footers/styles live elsewhere in the package.
 */
export interface DocumentBody {
  /** All content (paragraphs, tables) */
  content: BlockContent[];
  /** Sections (derived from sectPr in paragraphs and final sectPr) */
  sections?: Section[];
  /** Final section properties (from body's sectPr) */
  finalSectionProperties?: SectionProperties;
  /** Comments from comments.xml */
  comments?: Comment[];
  /** Stable reviewer palette metadata. Undefined = derive colors by author order. */
  commentAuthors?: CommentAuthor[];
}

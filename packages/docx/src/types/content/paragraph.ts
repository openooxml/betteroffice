/**
 * Paragraph (`w:p`) — the union of inline content that can sit inside a
 * paragraph (runs, hyperlinks, bookmarks, fields, SDT, comment ranges,
 * tracked-change wrappers, math) plus paragraph-level metadata
 * (formatting, list rendering, optional terminating section properties).
 */

import type { ParagraphFormatting } from '../formatting';
import type { ListRendering } from '../lists';
import type { Run } from './run';
import type { RawXml } from './rawXml';
import type { Hyperlink, BookmarkStart, BookmarkEnd, SimpleField, ComplexField } from './link';
import type { InlineSdt } from './sdt';
import type { CommentRangeStart, CommentRangeEnd } from './comment';
import type {
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
  MoveFromRangeStart,
  MoveFromRangeEnd,
  MoveToRangeStart,
  MoveToRangeEnd,
  ParagraphPropertyChange,
  TrackedChangeInfo,
} from './trackedChange';
import type { MathEquation } from './math';
import type { SectionProperties } from './section';

/**
 * Inline content that can appear inside a paragraph. Covers runs (text),
 * hyperlinks, bookmarks, fields, structured document tags, comment range
 * markers, tracked-change wrappers, and math equations. Every node in
 * this union carries a `type` discriminator so consumers can narrow at
 * runtime.
 */
export type ParagraphContent =
  | Run
  | Hyperlink
  | BookmarkStart
  | BookmarkEnd
  | SimpleField
  | ComplexField
  | InlineSdt
  | CommentRangeStart
  | CommentRangeEnd
  | Insertion
  | Deletion
  | MoveFrom
  | MoveTo
  | MoveFromRangeStart
  | MoveFromRangeEnd
  | MoveToRangeStart
  | MoveToRangeEnd
  | MathEquation
  | RawXml;

/**
 * Paragraph (`w:p`) — the primary block-level container in a Word document.
 *
 * Every paragraph carries direct formatting (`formatting`), tracked
 * property changes (`propertyChanges`), inline content (`content`), and
 * optional list rendering / section break metadata. `paraId` is Word's
 * paragraph identifier (`w14:paraId`); editing sessions address paragraphs
 * by their own session keys instead.
 *
 * See ECMA-376 §17.3.1.
 */
export interface Paragraph {
  type: 'paragraph';
  /**
   * Word paragraph ID (`w14:paraId`), when the paragraph has a valid one.
   * Invalid authored values are kept verbatim in `extraAttributes` but are
   * not an identity.
   */
  paraId?: string;
  /** Set when `paraId` repeats an ID used earlier in the package; it is still this paragraph's own. */
  repeatedParaId?: boolean;
  /**
   * The extra attribute holding an invalid Word 2010 `paraId` under a prefix
   * other than `w14`; a save that sets `paraId` replaces it.
   */
  paraIdAttribute?: string;
  /** This paragraph's `w:p` occurrence in its source part, on documents parsed for an editing session. */
  sourceOrdinal?: number;
  /** Text ID */
  textId?: string;
  /** `w:p` attributes the model does not type, kept verbatim so a save re-emits them. */
  extraAttributes?: { name: string; value: string }[];
  /** Paragraph formatting */
  formatting?: ParagraphFormatting;
  /** Paragraph-level tracked property changes (w:pPrChange) */
  propertyChanges?: ParagraphPropertyChange[];
  /**
   * Paragraph-mark insertion tracking (`<w:pPr><w:rPr><w:ins/>`). Set when
   * this paragraph's terminating pilcrow was added as a tracked change —
   * e.g., the user pressed Enter mid-paragraph in suggesting mode. Reject
   * joins this paragraph with the following one.
   */
  pPrIns?: TrackedChangeInfo;
  /**
   * Paragraph-mark deletion tracking (`<w:pPr><w:rPr><w:del/>`). Set when
   * this paragraph's terminating pilcrow was deleted as a tracked change —
   * e.g., the user pressed Backspace at the start of the next paragraph in
   * suggesting mode. Accept joins this paragraph with the following one.
   */
  pPrDel?: TrackedChangeInfo;
  /** Paragraph content */
  content: ParagraphContent[];
  /** Computed list rendering (if this is a list item) */
  listRendering?: ListRendering;
  /** Word's cached layout says this paragraph started on a new rendered page. */
  renderedPageBreakBefore?: boolean;
  /** Section properties (if this paragraph ends a section) */
  sectionProperties?: SectionProperties;
}

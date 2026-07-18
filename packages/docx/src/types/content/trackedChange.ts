/**
 * Tracked-changes model — insertion/deletion/move wrappers, range
 * markers, and per-element property-change wrappers (`w:rPrChange`,
 * `w:pPrChange`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`) plus
 * structural changes (row/cell insert/delete/merge).
 */

import type {
  TextFormatting,
  ParagraphFormatting,
  TableFormatting,
  TableRowFormatting,
  TableCellFormatting,
} from '../formatting';
import type { Run } from './run';
import type { Hyperlink } from './link';

/**
 * Tracked change metadata (w:ins, w:del attributes)
 */
export interface TrackedChangeInfo {
  /** Revision ID */
  id: number;
  /** Author who made the change */
  author: string;
  /** Date of the change */
  date?: string;
}

/**
 * Tracked-change attribute triple as it appears on PM node attrs
 * (`paragraph.pPrIns`, `tableRow.trIns`, etc). Mirrors `TrackedChangeInfo`
 * but with a `null` date (PM attr defaults) and a `revisionId` name that
 * matches OOXML's `w:id` more idiomatically on the editor side.
 *
 * Round-trip pairs with `TrackedChangeInfo` via
 * `{ id, author, date? } ↔ { revisionId, author, date | null }`.
 */
export interface RevisionInfo {
  revisionId: number;
  author: string;
  date: string | null;
}

/**
 * Tracked-cell marker — the OOXML `<w:cellIns>` / `<w:cellDel>` /
 * `<w:cellMerge>` shape attached to a `TableCell` PM node and surfaced
 * to the layout model and painter for visual rendering.
 *
 * `kind` matches the OOXML element name (ins / del / merge).
 */
export interface CellMarker {
  kind: 'ins' | 'del' | 'merge';
  info: RevisionInfo;
}

/**
 * Generic tracked property-change wrapper metadata (w:*PrChange)
 */
export interface PropertyChangeInfo extends TrackedChangeInfo {
  /** Optional revision session ID */
  rsid?: string;
}

/**
 * Insertion wrapper (w:ins) — runs inserted by tracked changes
 */
export interface Insertion {
  type: 'insertion';
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Inserted content */
  content: (Run | Hyperlink)[];
}

/**
 * Deletion wrapper (w:del) — runs deleted by tracked changes
 */
export interface Deletion {
  type: 'deletion';
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Deleted content */
  content: (Run | Hyperlink)[];
}

/**
 * Move-from wrapper (w:moveFrom) â€” content moved away from this position
 */
export interface MoveFrom {
  type: 'moveFrom';
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Moved content */
  content: (Run | Hyperlink)[];
}

/**
 * Move-to wrapper (w:moveTo) â€” content moved into this position
 */
export interface MoveTo {
  type: 'moveTo';
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /** Moved content */
  content: (Run | Hyperlink)[];
}

/**
 * Move-from range start marker (w:moveFromRangeStart) — ECMA-376 §17.13.5.22
 * Pairs with moveFromRangeEnd to delimit the source of a move in the document.
 */
export interface MoveFromRangeStart {
  type: 'moveFromRangeStart';
  id: number;
  name: string;
}

/**
 * Move-from range end marker (w:moveFromRangeEnd)
 */
export interface MoveFromRangeEnd {
  type: 'moveFromRangeEnd';
  id: number;
}

/**
 * Move-to range start marker (w:moveToRangeStart) — ECMA-376 §17.13.5.24
 * Pairs with moveToRangeEnd to delimit the destination of a move.
 */
export interface MoveToRangeStart {
  type: 'moveToRangeStart';
  id: number;
  name: string;
}

/**
 * Move-to range end marker (w:moveToRangeEnd)
 */
export interface MoveToRangeEnd {
  type: 'moveToRangeEnd';
  id: number;
}

/**
 * Run-level tracked wrappers represented in WordprocessingML.
 */
export type TrackedRunChange = Insertion | Deletion | MoveFrom | MoveTo;

/**
 * Run property change (w:rPrChange)
 */
export interface RunPropertyChange {
  type: 'runPropertyChange';
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Run properties before the tracked change */
  previousFormatting?: TextFormatting;
  /** Run properties after the tracked change (editor model convenience) */
  currentFormatting?: TextFormatting;
}

/**
 * Paragraph property change (w:pPrChange)
 */
export interface ParagraphPropertyChange {
  type: 'paragraphPropertyChange';
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Paragraph properties before the tracked change */
  previousFormatting?: ParagraphFormatting;
  /** Paragraph properties after the tracked change (editor model convenience) */
  currentFormatting?: ParagraphFormatting;
}

/**
 * Table property change (w:tblPrChange)
 */
export interface TablePropertyChange {
  type: 'tablePropertyChange';
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Table properties before the tracked change */
  previousFormatting?: TableFormatting;
  /** Table properties after the tracked change (editor model convenience) */
  currentFormatting?: TableFormatting;
}

/**
 * Table row property change (w:trPrChange)
 */
export interface TableRowPropertyChange {
  type: 'tableRowPropertyChange';
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Row properties before the tracked change */
  previousFormatting?: TableRowFormatting;
  /** Row properties after the tracked change (editor model convenience) */
  currentFormatting?: TableRowFormatting;
}

/**
 * Table cell property change (w:tcPrChange)
 */
export interface TableCellPropertyChange {
  type: 'tableCellPropertyChange';
  /** Tracked change metadata */
  info: PropertyChangeInfo;
  /** Cell properties before the tracked change */
  previousFormatting?: TableCellFormatting;
  /** Cell properties after the tracked change (editor model convenience) */
  currentFormatting?: TableCellFormatting;
}

/**
 * Table structural tracked change metadata (row/cell insert/delete/merge)
 */
export interface TableStructuralChangeInfo {
  type:
    | 'tableRowInsertion'
    | 'tableRowDeletion'
    | 'tableCellInsertion'
    | 'tableCellDeletion'
    | 'tableCellMerge';
  /** Tracked change metadata */
  info: TrackedChangeInfo;
  /**
   * `<w:cellMerge w:vMerge="…">` value, only meaningful for `tableCellMerge`.
   * Schema `ST_AnnotationVMerge`: `"rest"` = anchor (start of merged span),
   * `"cont"` = continuation (merged into predecessor). Word's default for a
   * tracked merge is `"cont"` (most edits track "this cell got merged INTO
   * the one above"); we preserve the on-disk value when present.
   */
  vMerge?: 'rest' | 'cont';
  /** `<w:cellMerge w:vMergeOrig="…">` — the pre-merge vMerge state. */
  vMergeOrig?: 'rest' | 'cont';
}

import type {
  ColorValue,
  ParagraphFormatting,
  TextFormatting,
} from '@betteroffice/docx/types/document';
import type {
  CollaborationPresence,
  CollaborationReplica,
  CollaborationUser,
} from '@betteroffice/docx/collaboration';
import type { PointPosition } from '@betteroffice/docx/plugin-api';
import type { DocxTextRange } from '@betteroffice/docx/yrs';

/**
 * The text under a client point in the rendered layout of `version`. `position`, `pageIndex`
 * and `region` locate the hit in that layout. `target` is the collapsed accepted-view range at
 * the point, keyed by session paragraph keys: an edit batch step can use it as its `target`
 * (`insertText` at `'start'` inserts at the point) with `expectVersion: version`.
 */
export interface DocxPointPosition extends PointPosition {
  version: string;
  target: { kind: 'range' } & DocxTextRange;
}

export interface DocxEditorCollaborationOptions {
  clientId?: number;
  user?: CollaborationUser;
  /** Shared Yrs state used instead of importing the source DOCX. */
  initialUpdate?: Uint8Array;
  onReplica?: (replica: CollaborationReplica | null) => void;
  presence?: CollaborationPresence;
}

/** Framework-neutral selection state published by the Yrs-backed editor. */
export interface SelectionState {
  hasSelection: boolean;
  isMultiParagraph: boolean;
  textFormatting: TextFormatting;
  paragraphFormatting: ParagraphFormatting;
  styleId: string | null;
  startParagraphIndex: number;
  endParagraphIndex: number;
}

/** Yrs-derived table context consumed by the toolbar. */
export interface TableContextInfo {
  isInTable: boolean;
  table?: { attrs?: { justification?: string } };
  rowIndex?: number;
  columnIndex?: number;
  rowCount?: number;
  columnCount?: number;
  hasMultiCellSelection?: boolean;
  canSplitCell?: boolean;
  cellBorderColor?: ColorValue;
  cellBackgroundColor?: string;
}

import type {
  ColorValue,
  ParagraphFormatting,
  TextFormatting,
} from '@betteroffice/docx/types/document';
import type { CollaborationReplica } from '@betteroffice/docx/collaboration';

export interface DocxEditorCollaborationOptions {
  clientId?: number;
  /** Shared Yrs state used instead of importing the source DOCX. */
  initialUpdate?: Uint8Array;
  onReplica?: (replica: CollaborationReplica | null) => void;
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

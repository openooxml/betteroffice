import type { TranslationKey } from '@betteroffice/docx-i18n';
import type { ImageLayoutTarget } from '@betteroffice/docx/docx';
import type { ParagraphAlignment } from '@betteroffice/docx/types/document';
import type {
  CommandReason,
  CommandState,
} from '../../../../shared/host-contracts/commands';
import type { EditorMode } from '../components/DocxEditor/internals/editing-modes';

export type {
  CommandReason,
  CommandState,
  JsonValue,
} from '../../../../shared/host-contracts/commands';

/** Table operations available from the table toolbar and menus. */
export type DocxTableAction =
  | 'addRowAbove'
  | 'addRowBelow'
  | 'addColumnLeft'
  | 'addColumnRight'
  | 'deleteRow'
  | 'deleteColumn'
  | 'mergeCells'
  | 'splitCell'
  | 'deleteTable'
  | 'selectTable'
  | 'selectRow'
  | 'selectColumn'
  | 'borderAll'
  | 'borderOutside'
  | 'borderInside'
  | 'borderNone'
  | 'borderTop'
  | 'borderBottom'
  | 'borderLeft'
  | 'borderRight'
  | { type: 'cellFillColor'; color: string | null }
  | { type: 'borderColor'; color: string }
  | { type: 'borderWidth'; size: number }
  | {
      type: 'cellBorder';
      side: 'top' | 'bottom' | 'left' | 'right' | 'all';
      style: string;
      size: number;
      color: string;
    }
  | { type: 'cellVerticalAlign'; align: 'top' | 'center' | 'bottom' }
  | {
      type: 'cellMargins';
      margins: { top?: number; bottom?: number; left?: number; right?: number };
    }
  | { type: 'cellTextDirection'; direction: string | null }
  | { type: 'toggleNoWrap' }
  | { type: 'rowHeight'; height: number | null; rule?: 'auto' | 'atLeast' | 'exact' }
  | { type: 'toggleHeaderRow' }
  | { type: 'distributeColumns' }
  | { type: 'autoFitContents' }
  | {
      type: 'tableProperties';
      props: {
        width?: number | null;
        widthType?: string | null;
        justification?: 'left' | 'center' | 'right' | null;
      };
    }
  | { type: 'openTableProperties' }
  | { type: 'applyTableStyle'; styleId: string };

/**
 * A direct text color: RGB hex, or a theme color optionally lightened
 * (`themeTint`) or darkened (`themeShade`) by a hex byte; `'auto'` clears it
 * to Word's automatic color.
 */
export type DocxTextColor =
  | { rgb: string }
  | { themeColor: string; themeTint?: string; themeShade?: string }
  | 'auto';

/** Image rotation and mirroring. */
export type DocxImageTransform = 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV';

/** Formatting of the table cell at the selection. */
export type DocxTableValue = {
  /** RGB hex without `#`. */
  borderColor: string | null;
  /** RGB hex without `#`. */
  fillColor: string | null;
  justification: string | null;
};

/** Arguments of every DOCX editor command, keyed by command id. */
export interface DocxCommandArgs {
  undo: null;
  redo: null;
  bold: null;
  italic: null;
  underline: null;
  strikethrough: null;
  superscript: null;
  subscript: null;
  clearFormatting: null;
  paragraphStyle: { styleId: string };
  fontFamily: { family: string };
  /** Size in points. */
  fontSize: { points: number };
  textColor: { color: DocxTextColor };
  /** A Word highlight name or RGB hex; `'none'` removes the highlight. */
  highlightColor: { color: string };
  alignment: { value: ParagraphAlignment };
  /** Auto line spacing in twips; 240 is single spacing. */
  lineSpacing: { value: number };
  bulletList: null;
  numberedList: null;
  indent: null;
  outdent: null;
  setLtr: null;
  setRtl: null;
  insertLink: null;
  insertImage: null;
  insertTable: { rows: number; columns: number };
  insertPageBreak: null;
  insertSectionBreakNextPage: null;
  insertSectionBreakContinuous: null;
  insertTOC: null;
  imageWrap: { wrap: ImageLayoutTarget };
  imageTransform: { action: DocxImageTransform };
  imageProperties: null;
  tableAction: DocxTableAction;
  pageSetup: null;
  watermark: null;
  editingMode: { mode: EditorMode };
  /** `null` resolves the tracked change at the selection. */
  reviewAccept: { revisionId: string } | null;
  /** `null` resolves the tracked change at the selection. */
  reviewReject: { revisionId: string } | null;
  reviewPrevious: null;
  reviewNext: null;
  commentsSidebar: null;
  open: null;
  save: null;
  print: null;
  find: null;
  replace: null;
  reportIssue: null;
  /** 1 is 100%. */
  zoom: { scale: number };
}

export type DocxCommandId = keyof DocxCommandArgs;

/** Commands presented as a choice between options. */
export type DocxSelectCommandId =
  | 'paragraphStyle'
  | 'fontFamily'
  | 'fontSize'
  | 'alignment'
  | 'lineSpacing'
  | 'imageWrap'
  | 'editingMode'
  | 'zoom';

/** Current values reported in {@link DocxCommandState.value}, keyed by command id. */
export interface DocxCommandValues {
  undo: null;
  redo: null;
  bold: null;
  italic: null;
  underline: null;
  strikethrough: null;
  superscript: null;
  subscript: null;
  clearFormatting: null;
  paragraphStyle: string;
  fontFamily: string | null;
  /** Points; `null` when mixed or unknown. */
  fontSize: number | null;
  /** RGB hex without `#`, or a theme color name. */
  textColor: string | null;
  highlightColor: string | null;
  alignment: ParagraphAlignment | null;
  lineSpacing: number | null;
  bulletList: null;
  numberedList: null;
  indent: null;
  outdent: null;
  setLtr: null;
  setRtl: null;
  insertLink: null;
  insertImage: null;
  insertTable: null;
  insertPageBreak: null;
  insertSectionBreakNextPage: null;
  insertSectionBreakContinuous: null;
  insertTOC: null;
  imageWrap: ImageLayoutTarget | null;
  imageTransform: null;
  imageProperties: null;
  tableAction: DocxTableValue | null;
  pageSetup: null;
  watermark: null;
  editingMode: EditorMode;
  reviewAccept: null;
  reviewReject: null;
  reviewPrevious: null;
  reviewNext: null;
  commentsSidebar: null;
  open: null;
  save: null;
  print: null;
  find: null;
  replace: null;
  reportIssue: null;
  zoom: number;
}

/** Why a command is unavailable. */
export type DocxCommandDisabledCode =
  | 'editor-unavailable'
  | 'document-loading'
  | 'no-document'
  | 'read-only'
  | 'viewing-mode'
  | 'suggesting-unsupported'
  | 'selection-required'
  | 'unsupported-selection'
  | 'unsupported-story'
  | 'cannot-outdent'
  | 'nothing-to-undo'
  | 'nothing-to-redo'
  | 'image-required'
  | 'table-required'
  | 'multiple-cells-required'
  | 'cannot-split-cell'
  | 'last-row'
  | 'last-column'
  | 'revision-required'
  | 'revision-not-found'
  | 'no-revisions'
  | 'controlled-mode'
  | 'controlled-sidebar'
  | 'host-disabled'
  | 'unsupported-command'
  | 'invalid-arguments';

/** Why an executed command did not complete. */
export type DocxCommandFailureCode =
  | DocxCommandDisabledCode
  | 'input-failed'
  | 'document-replaced'
  | 'target-changed'
  | 'command-failed';

/** Presentation hints for an option preview. */
export type DocxCommandOptionPreview = {
  /** CSS `font-family` value. */
  fontFamily?: string;
  /** Points. */
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  /** RGB hex without `#`. */
  color?: string;
  /** Grouping key, such as a font category or `document`. */
  group?: string;
};

/** One choice of a selector command. */
export interface DocxCommandOption<K extends DocxCommandId = DocxCommandId> {
  args: DocxCommandArgs[K];
  label: string;
  preview?: DocxCommandOptionPreview;
}

/** Serializable state of one DOCX command, optionally for specific arguments. */
export type DocxCommandState<K extends DocxCommandId = DocxCommandId> = CommandState<
  DocxCommandValues[K],
  DocxCommandDisabledCode
> & {
  /** Choices a selector presents, in display order. */
  options?: readonly DocxCommandOption<K>[];
};

/** A keyboard binding; `Mod` is Cmd on macOS and Ctrl elsewhere. */
export interface DocxCommandShortcut<K extends DocxCommandId = DocxCommandId> {
  chord: string;
  args: DocxCommandArgs[K];
}

/** Static, serializable description of a command. */
export interface DocxCommandDescriptor<K extends DocxCommandId = DocxCommandId> {
  id: K;
  labelKey: TranslationKey;
  mutatesDocument: boolean;
  shortcuts: readonly DocxCommandShortcut<K>[];
}

export type DocxCommandStatus = 'executed' | 'noop' | 'opened' | 'requested';

/**
 * Outcome of {@link DocxCommandStore.execute}: `executed` changed something,
 * `noop` was valid but changed nothing, `opened` showed a dialog or picker,
 * and `requested` handed the change to the host.
 */
export type DocxCommandResult =
  | { ok: true; status: DocxCommandStatus }
  | { ok: false; failure: CommandReason<DocxCommandFailureCode> };

/** The command authority of one editor, shared by built-in and host chrome. */
export interface DocxCommandStore {
  getDescriptor<K extends DocxCommandId>(id: K): DocxCommandDescriptor<K>;
  /** Snapshots are stable until the state changes; pass `args` to evaluate one option. */
  getState<K extends DocxCommandId>(id: K, args?: DocxCommandArgs[K]): DocxCommandState<K>;
  subscribe(listener: () => void): () => void;
  /** Runs after input accepted before the call; availability is checked again first. */
  execute<K extends DocxCommandId>(id: K, args: DocxCommandArgs[K]): Promise<DocxCommandResult>;
}

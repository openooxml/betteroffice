import type {
  Selection, WorkbookHandle, XlsxEditRequest, XlsxEditResult, XlsxFindRequest, XlsxFindResult,
  XlsxReadRequest, XlsxReadResult, XlsxValidationResult,
} from '@betteroffice/xlsx';
import type { CollaborationAwareness, CollaborationReplica } from '@betteroffice/xlsx/collaboration';
import type { Translations } from '@betteroffice/xlsx-i18n';
import type { XlsxCommandStore } from './commands/types';
import type { XlsxEditorPluginProps } from './plugins/types';
import { LEGACY_SYNC_API } from './runtimeOptions';
import type { XlsxAsyncEditorApi } from './worker/createAsyncEditorApi';
import type { XlsxWorkerEditorApi } from './worker/createWorkerEditorApi';

export interface XlsxPointPosition {
  sheet: number;
  row: number;
  col: number;
}

/**
 * The imperative surface handed to {@link XlsxEditorProps.onReady}: the open
 * workbook handle plus a `refreshProposals` to re-read the pending list after an
 * external caller (e.g. a demo agent) stages proposals on the same handle.
 *
 * The version, read and edit-batch methods run in order with the input accepted
 * before them, as commands do: cell and formula entries, chart moves, pastes and
 * cuts land first, and an IME composition ends with its text written. They reject
 * with {@link XlsxCommandAdmissionError}: `input-failed` while a refused entry
 * waits for correction or earlier input failed, `gesture-active` during a chart
 * drag, and `document-replaced` when the workbook is replaced meanwhile.
 * @deprecated Removed in 0.6.0. Use XlsxAsyncEditorApi.
 */
export interface XlsxEditorApi {
  flushPendingInput: () => Promise<void>;
  /** Client coordinates; sheet, row and column are zero-based. */
  getPositionAtPoint: (clientX: number, clientY: number) => XlsxPointPosition | null;
  /**
   * Closes the open cell entry and clears the selection. The entry is written
   * at once, or queued behind input still waiting to be written. An entry the
   * workbook refuses stays open, and nothing changes; one it refused earlier
   * on this sheet reopens at its own cell.
   */
  clearSelection: () => void;
  /**
   * The editor's command store, shared by its toolbar and host chrome. It
   * gates the editor's own UI; `handle` stays unrestricted host authority.
   * @experimental
   */
  readonly commands: XlsxCommandStore;
  focus: () => void;
  handle: WorkbookHandle;
  refreshProposals: () => void;
  /**
   * Writes the open cell entry and returns the workbook bytes at once. Throws
   * {@link XlsxSaveRefusedError} rather than return bytes without accepted
   * input: `input-pending` while earlier entries or a paste still wait to be
   * written (await `commands.execute('save', null)` instead), `input-failed`
   * while an entry the workbook refused waits for correction.
   */
  save: () => Uint8Array;
  /**
   * Scrolls the focus cell into view. Commands run after it, even in the same
   * handler, act on this selection and sheet. The open cell entry is written
   * first, or queued behind input still waiting to be written, so it may not
   * have landed when this returns. False, changing nothing, for an invalid
   * target or an open entry the workbook refuses; an entry it refused earlier
   * on `sheet` reopens at its own cell.
   */
  selectCells: (sheet: number, selection: Selection) => boolean;
  version: () => Promise<string>;
  readCells: (request: XlsxReadRequest) => Promise<XlsxReadResult>;
  findText: (request: XlsxFindRequest) => Promise<XlsxFindResult>;
  /** Refuses with `read-only` while the editor is read-only. */
  validateEdits: (request: XlsxEditRequest) => Promise<XlsxValidationResult>;
  /**
   * Applies the batch against the caller's `expectVersion` after committing pending input,
   * so input that lands first refuses it with `stale-version`. Refuses with `read-only`
   * while the editor is read-only; an applied batch calls `onChange` once.
   */
  applyEdits: (request: XlsxEditRequest) => Promise<XlsxEditResult>;
}

/** @deprecated Removed in 0.6.0. Use XlsxAsyncEditorApi. */
export interface XlsxWorkerViewerApi extends Omit<
  XlsxEditorApi, 'handle' | 'save' | 'selectCells' | 'version' | 'readCells' |
  'findText' | 'validateEdits' | 'applyEdits'
> {
  handle: null;
  save: () => null;
  selectCells: (sheet: number, selection: Selection) => false;
  saveAsync: () => Promise<Uint8Array | null>;
  selectCellsAsync: (sheet: number, selection: Selection) => Promise<boolean>;
  version: () => Promise<string | null>;
  readCells: (request: XlsxReadRequest) => Promise<XlsxReadResult | null>;
  findText: (request: XlsxFindRequest) => Promise<XlsxFindResult | null>;
  validateEdits: (request: XlsxEditRequest) => Promise<XlsxValidationResult | null>;
  applyEdits: (request: XlsxEditRequest) => Promise<XlsxEditResult | null>;
}

export interface XlsxEditorCollaborationOptions {
  /** Peer-unique Yrs client ID. Generated securely when omitted. */
  clientId?: number;
  /** Shared Yrs state applied before the replica is exposed. */
  initialUpdate?: Uint8Array;
  /** Receive the editor-owned collaboration replica. */
  onReplica?: (replica: CollaborationReplica | null) => void;
  /** Presence-capable provider connected to the editor-owned replica. */
  provider?: CollaborationAwareness | null;
}

export interface XlsxDocumentChange { version: string }

interface XlsxHostProps {
  /** Receives the version after a committed edit changes the workbook. */
  onDocumentChange?: (change: XlsxDocumentChange) => void;
  /** Whether Save also downloads a copy. Defaults to true without onSave, false with it. */
  downloadOnSave?: boolean;
  /** Initial scale, clamped to the supported zoom range. Defaults to 1. */
  initialZoom?: number;
  /** Rendered while a supplied file is opening. */
  loadingIndicator?: React.ReactNode;
  /** Inline style overrides for the editor root. */
  style?: React.CSSProperties;
}

/** Props for XlsxEditor. */
export interface XlsxInThreadEditorProps extends XlsxEditorPluginProps, XlsxHostProps {
  /** @deprecated False selects the in-thread engine. */
  experimentalWorkerOpen?: false;
  onError?: (error: Error) => void;
  /** Raw .xlsx bytes to open. When omitted the shell paints a demo frame. */
  file?: Uint8Array | ArrayBuffer;
  /** Download name for the save button; falls back to `workbook.xlsx`. */
  fileName?: string;
  /** Receives saved workbook bytes. */
  onSave?: (bytes: Uint8Array) => void;
  /** Return true for built-in saving; false or void handles/cancels the request. */
  onSaveRequest?: () => boolean | void | Promise<boolean | void>;
  /** Called after a user edit changes the workbook. */
  onChange?: () => void;
  /** Open a network-ready Yrs replica and repaint when peer updates arrive. */
  collaboration?: XlsxEditorCollaborationOptions;
  i18n?: Translations;
  /**
   * Called when a workbook opens, with a handle to stage agent proposals and a
   * way to refresh the panel afterward. Enables demo/host agents without
   * exposing the wasm object through the render tree. A returned cleanup runs
   * before the workbook is replaced or disposed.
   */
  onReady?: (api: XlsxEditorApi) => void | (() => void);
  className?: string;
  /** Blocks user edits; navigation and selection remain available. */
  readOnly?: boolean;
  /**
   * Replaces the default toolbar: omitted keeps it (hidden when read-only),
   * `null` removes it, and supplied chrome renders instead, also when
   * read-only. Compose it from `EditorToolbar mode="commands"` parts.
   * @experimental
   */
  toolbar?: React.ReactNode;
  /** `false` hides the toolbar region, whatever `toolbar` is. */
  showToolbar?: boolean;
}

/** @deprecated Removed in 0.6.0. Use XlsxAsyncEditorProps. */
export type XlsxWorkerViewerProps = Omit<XlsxInThreadEditorProps, 'onReady' | 'readOnly' | 'collaboration' | 'experimentalWorkerOpen' | keyof XlsxHostProps> & {
  /** @deprecated Removed in 0.6.0. Omit for the async editor. */
  [LEGACY_SYNC_API]: true;
  readOnly: true;
  collaboration?: undefined;
  /** Opens in the worker; without it the legacy runtime runs in-thread. */
  experimentalWorkerOpen: true;
  onReady?: (api: XlsxWorkerViewerApi) => void | (() => void);
};

/** @deprecated Removed in 0.6.0. Use XlsxAsyncEditorProps. */
export type XlsxWorkerEditorProps = Omit<XlsxWorkerViewerProps, 'onReady' | 'readOnly'> & {
  readOnly?: false;
  onReady?: (api: XlsxWorkerEditorApi) => void | (() => void);
};

export type XlsxAsyncEditorProps = Omit<XlsxInThreadEditorProps, 'onReady' | 'experimentalWorkerOpen'> & {
  /** @deprecated Removed in 0.6.0. Omit for the async editor. */
  [LEGACY_SYNC_API]?: false;
  /** @deprecated False selects the in-thread engine. */
  experimentalWorkerOpen?: boolean;
  onReady?: (api: XlsxAsyncEditorApi) => void | (() => void);
};

/** @deprecated Removed in 0.6.0. Use XlsxAsyncEditorProps. */
export type XlsxLegacyEditorProps = Omit<XlsxInThreadEditorProps, keyof XlsxHostProps> & {
  /** @deprecated Removed in 0.6.0. Omit for the async editor. */
  [LEGACY_SYNC_API]: true;
};

export type WorkerSessionProps = Omit<XlsxWorkerEditorProps, 'readOnly' | 'onReady'> & {
  readOnly?: boolean;
  onReady?: (api: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => void | (() => void);
};
export type GeneralEditorSessionProps = Omit<XlsxLegacyEditorProps, 'experimentalWorkerOpen' | 'onReady'> & {
  experimentalWorkerOpen?: boolean;
  onReady?: (api: XlsxEditorApi | XlsxWorkerEditorApi | XlsxWorkerViewerApi) => void | (() => void);
};

export type XlsxEditorProps = XlsxAsyncEditorProps | GeneralEditorSessionProps;

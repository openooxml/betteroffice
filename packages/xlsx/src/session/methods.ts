import type { MethodPolicies, SessionState } from '../../../../shared/office-session';
import type { Direction, SelectionData } from '../selection/types';
import type { DisplayList, Rect } from '../display-list/types';
import type { CollaborationUpdateOrigin } from '../collaboration/types';
import type { XlsxExportOptions, XlsxExportResult, XlsxMarkdownContent, XlsxMarkdownOptions, XlsxStructuredContent } from '../exports';
import type {
  WorkbookCommand, WorkbookCommandContext, WorkbookCommandFrame, WorkbookCommandSelection,
  WorkbookCommandWireFrame, WorkbookEditState, WorkbookWireAck,
  WorkbookSharedCommand, WorkbookSharedWireAck, WorkbookCommandFrameRequest,
} from './commandProtocol';
import type { Checkpoint, FrameEnvelope } from '../../../../shared/office-session/commands/types';
import type {
  XlsxEditRequest,
  XlsxEditResult,
  XlsxFindRequest,
  XlsxFindResult,
  XlsxReadRequest,
  XlsxReadResult,
  XlsxValidationResult,
} from '../edits';
import type {
  CalculationStatus, CapturedFormat, CellEdit, CellPosition, MergedRange, OpenWorkbookOptions, PrintMetrics, SheetInfo, Viewport,
} from '../wasm/loader';

/** @experimental */
export class WorkbookPngUnavailableError extends Error {
  readonly code = 'png-unavailable';
  constructor() {
    super('PNG export is unavailable');
    this.name = 'WorkbookPngUnavailableError';
  }
}

export type WorkbookCheckpoint = Checkpoint & { encodeMs?: number };

/** @experimental Unzoomed sheet pixels. */
export type WorkbookSheetView = Pick<SheetInfo,
  'contentWidth' | 'contentHeight' | 'frozenRows' | 'frozenCols' | 'initialScrollX' | 'initialScrollY'
> & { sheet: number; version: string; frozenWidth: number; frozenHeight: number };

/** @experimental Absolute sheet rect and freeze-relative scroll position, in unzoomed pixels. */
export interface WorkbookCellGeometry {
  sheet: number;
  version: string;
  rect: Rect;
  scrollPosition: CellPosition;
}

/** @experimental Row-major editable cell text. */
export interface WorkbookCellInputs {
  sheet: number;
  version: string;
  cells: CellEdit[][];
}

/** @experimental */
export interface WorkbookFrameOptions {
  sheet?: number;
}

/** @experimental */
export interface WorkbookFrame {
  displayList: DisplayList;
  version: string;
  epoch: number;
  sequence: number;
  sheet: number;
  viewport: Viewport;
  mergedRanges?: MergedRange[];
}

/** @experimental */
export type WorkbookWireFrame = Omit<WorkbookFrame, 'displayList'> & {
  displayList: ArrayBuffer;
};

/**
 * Structured-cloneable workbook open options.
 * @experimental
 */
export interface WorkbookSessionOpenOptions extends OpenWorkbookOptions {
  wasm?: ArrayBuffer | WebAssembly.Module;
  initialViewport?: WorkbookInitialPaintRequest;
  includeEditState?: boolean;
  initialUpdate?: Uint8Array;
  /** @internal */
  recoveryVersion?: string;
}

/** @experimental */
export interface WorkbookInitialPaintRequest {
  sheet?: number;
  width: number;
  height: number;
  zoom: number;
  dpr: number;
  requestGeneration: number;
}

/** @experimental */
export interface WorkbookInitialPaint<F = WorkbookCommandFrame> {
  sheetView: WorkbookSheetView;
  frame: F;
  editState: WorkbookEditState;
}

/** @experimental */
export type WorkbookSessionOpenReply = WorkbookSessionState & {
  pngExport?: boolean;
  clientId?: number;
  throughSeq?: number;
  revision?: number;
  historyReset?: boolean;
  initialPaint?: WorkbookInitialPaint<WorkbookCommandWireFrame>;
  editState?: WorkbookEditState;
  selection?: WorkbookCommandSelection;
};

/**
 * Sheet metadata; indices are zero-based.
 * @experimental
 */
export interface WorkbookSheetSummary {
  id: string;
  index: number;
  name: string;
}

/**
 * Workbook projection; version counts applied batches.
 * @experimental
 */
export type WorkbookSessionState = Omit<SessionState, 'stage'> & {
  format: 'xlsx';
  stage: 'preview' | 'ready' | 'failed';
  sheets: WorkbookSheetSummary[];
  activeSheet: number;
  commandContext?: WorkbookCommandContext;
};

/**
 * Workbook change notifications after applied batches.
 * @experimental
 */
export interface WorkbookSessionEvents {
  changed: { version: number; dirty: boolean };
  collaborationUpdate: { update: Uint8Array; origin: CollaborationUpdateOrigin };
}

/**
 * Workbook RPC methods; engine version tokens remain strings.
 * @experimental
 */
export type WorkbookSessionMethods = {
  open(bytes: ArrayBuffer, options?: WorkbookSessionOpenOptions): WorkbookSessionOpenReply;
  openRecovery(checkpoint: Checkpoint, options?: Pick<WorkbookSessionOpenOptions, 'wasm'>): WorkbookSessionOpenReply;
  version(): string;
  collaborationStateVector(): Uint8Array;
  collaborationStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  readCells(request: XlsxReadRequest): XlsxReadResult;
  findText(request: XlsxFindRequest): XlsxFindResult;
  validateEdits(request: XlsxEditRequest): XlsxValidationResult;
  applyEdits(request: XlsxEditRequest): XlsxEditResult;
  command(command: WorkbookCommand): WorkbookWireAck;
  edit(command: WorkbookSharedCommand): WorkbookSharedWireAck;
  checkpoint(preserveHistory?: boolean): WorkbookCheckpoint;
  commandFrame(request?: WorkbookCommandFrameRequest): FrameEnvelope<WorkbookCommandWireFrame> | undefined;
  frame(viewport: Viewport, options?: WorkbookFrameOptions): WorkbookWireFrame;
  sheetView(sheet: number): WorkbookSheetView;
  cellGeometry(sheet: number, row: number, col: number): WorkbookCellGeometry;
  cellInputs(sheet: number, range: string): WorkbookCellInputs;
  selectionData(sheet: number, row: number, col: number, direction?: Direction): SelectionData;
  captureFormat(sheet: number, range: string): CapturedFormat;
  sheets(): WorkbookSheetSummary[];
  calculationStatus(): CalculationStatus;
  renderPng(viewport: Viewport, options?: WorkbookFrameOptions): ArrayBuffer;
  renderRangePng(options?: { range?: string; scale?: number }, frameOptions?: WorkbookFrameOptions): ArrayBuffer;
  exportStructured(options?: XlsxExportOptions): XlsxExportResult<XlsxStructuredContent>;
  exportMarkdown(options?: XlsxExportOptions, markdownOptions?: XlsxMarkdownOptions): XlsxExportResult<XlsxMarkdownContent>;
  printDisplayList(sheet: number, range: string, metrics: PrintMetrics, gridlines: boolean): ArrayBuffer;
  save(): ArrayBuffer;
  dispose(): void;
};

export const WORKBOOK_SESSION_METHODS = {
  open: true,
  openRecovery: true,
  version: true,
  collaborationStateVector: true,
  collaborationStateAsUpdate: true,
  readCells: true,
  findText: true,
  validateEdits: true,
  applyEdits: true,
  command: true,
  edit: true,
  checkpoint: true,
  commandFrame: true,
  frame: true,
  sheetView: true,
  cellGeometry: true,
  cellInputs: true,
  selectionData: true,
  captureFormat: true,
  sheets: true,
  calculationStatus: true,
  renderPng: true,
  renderRangePng: true,
  exportStructured: true,
  exportMarkdown: true,
  printDisplayList: true,
  save: true,
  dispose: true,
} satisfies { readonly [K in keyof WorkbookSessionMethods]-?: true };

export const WORKBOOK_SESSION_POLICIES: MethodPolicies<WorkbookSessionMethods> = {
  open: { lane: 'interactive' },
  openRecovery: { lane: 'input', reorderable: false },
  version: { lane: 'interactive' },
  collaborationStateVector: { lane: 'collab', reorderable: false },
  collaborationStateAsUpdate: { lane: 'collab', reorderable: false },
  readCells: { lane: 'interactive' },
  findText: { lane: 'interactive' },
  validateEdits: { lane: 'interactive' },
  applyEdits: { lane: 'input', mutates: true, userInput: true },
  command: { lane: 'input', userInput: true, reorderable: false },
  edit: { lane: 'input', userInput: true, mutates: true, reframes: true, reorderable: false },
  checkpoint: { lane: 'input', reorderable: false },
  commandFrame: { lane: 'input', reorderable: false },
  frame: { lane: 'interactive', reframes: true, key: 'frame', replaceableBy: 'frame' },
  sheetView: { lane: 'interactive' },
  cellGeometry: { lane: 'interactive' },
  cellInputs: { lane: 'interactive' },
  selectionData: { lane: 'interactive', reorderable: false },
  captureFormat: { lane: 'interactive', reorderable: false },
  sheets: { lane: 'interactive' },
  calculationStatus: { lane: 'interactive' },
  renderPng: { lane: 'interactive', reorderable: false },
  renderRangePng: { lane: 'interactive', reorderable: false },
  exportStructured: { lane: 'interactive', reorderable: false },
  exportMarkdown: { lane: 'interactive', reorderable: false },
  printDisplayList: { lane: 'interactive', reorderable: false },
  save: { lane: 'interactive' },
  dispose: { lane: 'interactive' },
};

import type { MethodPolicies, SessionState } from '../../../../shared/office-session';
import type { DisplayList } from '../display-list/types';
import type {
  XlsxEditRequest,
  XlsxEditResult,
  XlsxFindRequest,
  XlsxFindResult,
  XlsxReadRequest,
  XlsxReadResult,
  XlsxValidationResult,
} from '../edits';
import type { CalculationStatus, OpenWorkbookOptions, Viewport } from '../wasm/loader';

/** @experimental */
export interface WorkbookFrameOptions {
  sheet?: number;
}

/** @experimental */
export interface WorkbookFrame {
  displayList: DisplayList;
  version: string;
  epoch: number;
  sheet: number;
  viewport: Viewport;
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
}

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
};

/**
 * Workbook change notifications after applied batches.
 * @experimental
 */
export interface WorkbookSessionEvents {
  changed: { version: number; dirty: boolean };
}

/**
 * Workbook RPC methods; engine version tokens remain strings.
 * @experimental
 */
export type WorkbookSessionMethods = {
  open(bytes: ArrayBuffer, options?: WorkbookSessionOpenOptions): WorkbookSessionState;
  version(): string;
  readCells(request: XlsxReadRequest): XlsxReadResult;
  findText(request: XlsxFindRequest): XlsxFindResult;
  validateEdits(request: XlsxEditRequest): XlsxValidationResult;
  applyEdits(request: XlsxEditRequest): XlsxEditResult;
  frame(viewport: Viewport, options?: WorkbookFrameOptions): WorkbookWireFrame;
  sheets(): WorkbookSheetSummary[];
  calculationStatus(): CalculationStatus;
  save(): ArrayBuffer;
  dispose(): void;
};

export const WORKBOOK_SESSION_METHODS = {
  open: true,
  version: true,
  readCells: true,
  findText: true,
  validateEdits: true,
  applyEdits: true,
  frame: true,
  sheets: true,
  calculationStatus: true,
  save: true,
  dispose: true,
} satisfies { readonly [K in keyof WorkbookSessionMethods]-?: true };

export const WORKBOOK_SESSION_POLICIES: MethodPolicies<WorkbookSessionMethods> = {
  open: { lane: 'interactive' },
  version: { lane: 'interactive' },
  readCells: { lane: 'interactive' },
  findText: { lane: 'interactive' },
  validateEdits: { lane: 'interactive' },
  applyEdits: { lane: 'input', mutates: true, userInput: true },
  frame: { lane: 'interactive', reframes: true, key: 'frame', replaceableBy: 'frame' },
  sheets: { lane: 'interactive' },
  calculationStatus: { lane: 'interactive' },
  save: { lane: 'interactive' },
  dispose: { lane: 'interactive' },
};

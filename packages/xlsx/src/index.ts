/**
 * `@betteroffice/xlsx` — the framework-free core: display-list types, the Canvas2D
 * backend, pure viewport math, and the wasm loader. Framework chrome lives in
 * the adapter packages (`@betteroffice/xlsx-react`); nothing here touches react/vue
 * (lint-enforced).
 */

export * from './worker';
/** @internal */
export { hydratePeer } from './session/client';
/** @experimental */
export { WorkbookPeerHydrationError } from './session/peerHydrationError';
/** @internal */
export { createWorkbookEditPeer } from './session/editPeer';
/** @internal */
export { localNowSerial } from './session/calculationClock';
/** @internal */
export type { WorkbookEditPeer, WorkbookEditPeerOptions } from './session/editPeer';
/** @experimental */
export { WorkbookEditPeerFailedError } from './session/editPeer';
/** @internal */
export {
  createWorkbookRecoveryMutators, failWorkbookEditPeer, WorkbookRecoveryRefusal, workbookEditPeerOperations,
} from './session/editPeerInternals';
/** @internal */
export { workbookSessionInternals } from './session/replay';
/** @internal */
export type { WorkbookReplayOp } from './session/replay';

export type {
  Rect,
  TextAlign,
  FillRectCmd,
  PatternFillCmd,
  PatternType,
  LineCmd,
  GeometryPathCommand,
  PathStroke,
  PathCmd,
  TextCmd,
  DrawCmd,
  DisplayList,
  GridMeta,
  HyperlinkRegion,
  ChartRegion,
  ChartA11yAttrs,
} from './display-list/types';

export type { TrackOffsets, ViewportState, VisibleRange, VisibleCells } from './viewport/index';

export {
  trackCount,
  totalExtent,
  visibleRange,
  visibleCells,
  clampScroll,
  uniformOffsets,
} from './viewport/index';

export type {
  CellAddr,
  Selection,
  CellRange,
  Direction,
  SelectionLimits,
  SelectionData,
  SelectionHidden,
  SelectionMerges,
  MoveOpts,
  KeyInput,
  SelectionAction,
} from './selection/index';

export {
  selectionAt,
  normalizeRange,
  rangeContains,
  cellCount,
  extendTo,
  moveFocus,
  selectionKeyReducer,
} from './selection/index';

export { cellAtPoint, cellRect, chartRegionAtPoint, rangeRect } from './hittest/index';

export type { HyperlinkDestination } from './hyperlinks/index';
export {
  hyperlinkAtCell,
  parseHyperlinkLocation,
  safeExternalHyperlink,
} from './hyperlinks/index';

export type { CellInput } from './clipboard/index';
export { toTsv, fromTsv } from './clipboard/index';

export type {
  A11yStrings,
  A11yCell,
  A11yChart,
  A11yColumnHeader,
  A11yRow,
  A11yGrid,
} from './a11y/index';
export { buildA11yGrid } from './a11y/index';

export { paintDisplayList } from './render/canvas2d';
export {
  initWasm,
  isWasmAvailable,
  isPngExportAvailable,
  isProposalsAvailable,
  openWorkbook,
  wasmVersion,
  StaleProposalError,
  DisplayTooLargeError,
  getDisplayListCellLimit,
  exportXlsxMarkdown,
  exportXlsxStructured,
  renderXlsxMarkdown,
} from './wasm/loader';
export type {
  WasmInitInput,
  OpenWorkbookOptions,
  WorkbookCalculationContext,
  Viewport,
  PrintMetrics,
  SheetInfo,
  WorkbookHandle,
  XlsxTextMatch,
  XlsxTextSearchOptions,
  WorkbookUpdateListener,
  WorkbookUpdateOrigin,
  CellEdit,
  CellPosition,
  CellInputEdit,
  EditResult,
  EditProfile,
  ProfiledEditResult,
  DisplayListProfile,
  ProfiledDisplayList,
  CalculationStatus,
  BorderPatch,
  BorderPreset,
  BorderStyle,
  CapturedFormat,
  CapturedColor,
  CapturedFill,
  CapturedBorderStyle,
  CapturedBorderEdge,
  CapturedCellFormat,
  CellPoint,
  ConditionalFormatInfo,
  XlsxAutoFilterOp,
  ConditionalFormatOp,
  ConditionalFormatRuleInput,
  DifferentialFormatInput,
  HistoryState,
  HorizontalAlignment,
  MergedRange,
  WorkbookHyperlink,
  NumberFormat,
  NumberFormatMutation,
  Proposal,
  ProposalCell,
  ProposalEdit,
  RangeStylePatch,
  SelectionFormatting,
  StaleProposalTarget,
  StyleProperty,
  TextWrapping,
  VerticalAlignment,
} from './wasm/loader';
export type {
  EditSuccess,
  OperationFailure,
  OperationRefusal,
  ValidationSuccess,
  XlsxCellAddress,
  XlsxCellGuard,
  XlsxCellPosition,
  XlsxCellRead,
  XlsxCellValue,
  XlsxEditCalculation,
  XlsxEditFailure,
  XlsxEditFailureCode,
  XlsxEditHistory,
  XlsxEditOperation,
  XlsxEditPreview,
  XlsxEditReceipt,
  XlsxEditRefusal,
  XlsxEditRequest,
  XlsxEditResult,
  XlsxEditSource,
  XlsxEditStep,
  XlsxErrorValue,
  XlsxFindMatch,
  XlsxFindRequest,
  XlsxFindResult,
  XlsxRangeAddress,
  XlsxRangeRead,
  XlsxRangeTarget,
  XlsxReadCalculation,
  XlsxReadRequest,
  XlsxReadResult,
  XlsxSheetEntry,
  XlsxValidationResult,
} from './edits';
export type {
  ExportCompletion,
  ExportDiagnostic,
  ExportSeverity,
  MarkdownAnchor,
  XlsxAnchor,
  XlsxExportCell,
  XlsxExportDefinedName,
  XlsxExportDiagnostic,
  XlsxExportDiagnosticCode,
  XlsxExportFailure,
  XlsxExportFailureCode,
  XlsxExportHyperlink,
  XlsxExportMerge,
  XlsxExportObject,
  XlsxExportOptions,
  XlsxExportResult,
  XlsxExportScope,
  XlsxExportSheet,
  XlsxExportTable,
  XlsxExportValue,
  XlsxFormulaResult,
  XlsxMarkdownContent,
  XlsxMarkdownOptions,
  XlsxObjectKind,
  XlsxSheetIdentity,
  XlsxSourcePart,
  XlsxStructuredContent,
} from './exports';

export type { XlsxDefinedNameOp, XlsxDefinedName, XlsxDefinedNamesRequest, XlsxDefinedNamesResult } from './definedNames';

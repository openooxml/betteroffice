/**
 * `@betteroffice/xlsx` — the framework-free core: display-list types, the Canvas2D
 * backend, pure viewport math, and the wasm loader. Framework chrome lives in
 * the adapter packages (`@betteroffice/xlsx-react`); nothing here touches react/vue
 * (lint-enforced).
 */

export type {
  Rect,
  TextAlign,
  FillRectCmd,
  LineCmd,
  TextCmd,
  DrawCmd,
  DisplayList,
  GridMeta,
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

export { cellAtPoint, cellRect, rangeRect } from './hittest/index';

export type { CellInput } from './clipboard/index';
export { toTsv, fromTsv } from './clipboard/index';

export type { A11yStrings, A11yCell, A11yColumnHeader, A11yRow, A11yGrid } from './a11y/index';
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
} from './wasm/loader';
export type {
  WasmInitInput,
  OpenWorkbookOptions,
  Viewport,
  SheetInfo,
  WorkbookHandle,
  WorkbookUpdateListener,
  WorkbookUpdateOrigin,
  CellEdit,
  CellInputEdit,
  EditResult,
  CalculationStatus,
  Proposal,
  ProposalCell,
  ProposalEdit,
} from './wasm/loader';

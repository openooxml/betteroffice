/**
 * Headless entry — the DOM-free compute surface for server rendering, tests,
 * and workers. Re-exports only pure viewport math and the display-list types;
 * no canvas, no wasm loader, no framework. The seam-purity test walks the
 * import closure from this file and fails if anything impure creeps in.
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

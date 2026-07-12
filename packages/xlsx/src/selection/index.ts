/**
 * Pure selection seam: the anchor/focus model, its transforms, and the keyboard
 * reducer. DOM-free by construction (see the seam-purity test).
 */

export type {
  CellAddr,
  Selection,
  CellRange,
  Direction,
  SelectionLimits,
  MoveOpts,
  KeyInput,
  SelectionAction,
} from './types';

export {
  selectionAt,
  normalizeRange,
  rangeContains,
  cellCount,
  extendTo,
  moveFocus,
} from './model';

export { selectionKeyReducer } from './keyboard';

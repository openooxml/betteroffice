/**
 * Keyboard reducer: maps a keystroke over the current selection to a
 * declarative {@link SelectionAction}. Pure — it neither mutates state nor
 * touches the DOM; the chrome owns the keydown listener and interprets the
 * action. Excel navigation semantics, with two deliberate choices documented at
 * the F2/Enter handlers below.
 */

import { extendTo, moveFocus, normalizeRange, selectionAt } from './model';
import type {
  CellAddr,
  Direction,
  KeyInput,
  Selection,
  SelectionAction,
  SelectionLimits,
  SelectionData,
  SelectionHidden,
} from './types';

const ARROWS: Record<string, Direction> = {
  ArrowUp: 'up',
  ArrowDown: 'down',
  ArrowLeft: 'left',
  ArrowRight: 'right',
};

function mod(key: KeyInput): boolean {
  return Boolean(key.metaKey || key.ctrlKey);
}

function edgeInDirection(focus: CellAddr, direction: Direction, limits: SelectionLimits): CellAddr {
  switch (direction) {
    case 'up':
      return { row: 0, col: focus.col };
    case 'down':
      return { row: limits.rows - 1, col: focus.col };
    case 'left':
      return { row: focus.row, col: 0 };
    case 'right':
      return { row: focus.row, col: limits.cols - 1 };
  }
}

function moveAction(selection: Selection): SelectionAction {
  return { type: 'move', selection };
}

function visibleIndex(start: number, step: number, limit: number, hidden?: (index: number) => boolean): number | null {
  let index = start;
  while (index >= 0 && index < limit && hidden?.(index)) index += step;
  return index >= 0 && index < limit ? index : null;
}

function tabMove(sel: Selection, back: boolean, limits: SelectionLimits, hidden?: SelectionHidden): Selection {
  const { row, col } = sel.focus;
  const step = back ? -1 : 1;
  const next = visibleIndex(col + step, step, limits.cols, hidden?.isColHidden);
  if (next !== null) return selectionAt({ row, col: next });
  const wrappedCol = visibleIndex(back ? limits.cols - 1 : 0, step, limits.cols, hidden?.isColHidden);
  const wrappedRow = visibleIndex(row + step, step, limits.rows, hidden?.isRowHidden);
  if (wrappedCol === null || (hidden && wrappedRow === null)) return sel;
  return selectionAt({ row: wrappedRow ?? row, col: wrappedCol });
}

function visibleEdge(state: Selection, edge: CellAddr, direction: Direction, limits: SelectionLimits,
  hidden?: SelectionHidden): CellAddr {
  if (!hidden || (!hidden.isRowHidden(edge.row) && !hidden.isColHidden(edge.col))) return edge;
  const vertical = direction === 'up' || direction === 'down';
  const back = direction === 'up' || direction === 'left';
  const origin = vertical ? state.focus.row : state.focus.col;
  const target = vertical ? edge.row : edge.col;
  const limit = vertical ? limits.rows : limits.cols;
  const isHidden = vertical ? hidden.isRowHidden : hidden.isColHidden;
  const index = visibleIndex(target, back ? -1 : 1, limit, isHidden) ??
    visibleIndex(target, back ? 1 : -1, limit, isHidden);
  if (index === null || (back ? index >= origin : index <= origin)) return state.focus;
  return vertical ? { row: index, col: edge.col } : { row: edge.row, col: index };
}

function home(state: Selection, sheet: boolean, limits: SelectionLimits, hidden?: SelectionHidden): Selection {
  const col = visibleIndex(0, 1, limits.cols, hidden?.isColHidden);
  const row = sheet ? visibleIndex(0, 1, limits.rows, hidden?.isRowHidden) : state.focus.row;
  if (row === null || col === null || (hidden && row === state.focus.row && col === state.focus.col)) return state;
  return selectionAt({ row, col });
}

// a lone printable character (no ctrl/cmd/alt) opens the editor seeded with it.
function isPrintable(key: KeyInput): boolean {
  return key.key.length === 1 && !key.ctrlKey && !key.metaKey && !key.altKey;
}

/**
 * Reduce a keystroke to a selection action against the current selection and
 * grid limits. Returns `{ type: 'none' }` for keys this layer does not own so
 * the chrome can fall through to its own handling.
 */
export function selectionKeyReducer(
  state: Selection,
  key: KeyInput,
  limits: SelectionLimits,
  data?: SelectionData,
  hidden?: SelectionHidden
): SelectionAction {
  const direction = ARROWS[key.key];
  if (direction) {
    if (mod(key)) {
      const edge = visibleEdge(state, data?.edge ?? edgeInDirection(state.focus, direction, limits), direction, limits, hidden);
      if (hidden && edge.row === state.focus.row && edge.col === state.focus.col) return moveAction(state);
      return moveAction(key.shiftKey ? extendTo(state, edge) : selectionAt(edge));
    }
    return moveAction(moveFocus(state, direction, { extend: key.shiftKey, limits, hidden }));
  }

  switch (key.key) {
    case 'Tab':
      return moveAction(tabMove(state, Boolean(key.shiftKey), limits, hidden));

    // enter navigates (down/up); it never opens the editor — F2 does. this is
    // the excel split: type to overwrite, F2 to edit in place, enter to commit.
    case 'Enter':
      return moveAction(moveFocus(state, key.shiftKey ? 'up' : 'down', { extend: false, limits, hidden }));

    case 'Home':
      return moveAction(home(state, mod(key), limits, hidden));

    case 'End': {
      const used = data?.usedRange;
      const found = mod(key) ? used && { row: used.bottom, col: used.right } : data?.edge;
      if (found === undefined) return { type: 'none' };
      if (found === null && !mod(key)) return moveAction(state);
      const target = {
        row: Math.max(0, Math.min(found?.row ?? 0, limits.rows - 1)),
        col: Math.max(0, Math.min(found?.col ?? 0, limits.cols - 1)),
      };
      return moveAction(key.shiftKey ? extendTo(state, target, limits) : selectionAt(target));
    }

    case 'PageDown':
      return moveAction(
        moveFocus(state, 'down', { extend: key.shiftKey, step: limits.rowsPerPage, limits, hidden })
      );

    case 'PageUp':
      return moveAction(
        moveFocus(state, 'up', { extend: key.shiftKey, step: limits.rowsPerPage, limits, hidden })
      );

    // F2 opens the editor on the current cell without seeding input.
    case 'F2':
      return { type: 'startEdit' };

    case 'Delete':
    case 'Backspace':
      return { type: 'clear' };
  }

  if ((key.key === 'a' || key.key === 'A') && mod(key)) {
    if (data) {
      const current = normalizeRange(state);
      const region = data.region;
      const range = region && (current.top !== region.top || current.left !== region.left ||
        current.bottom !== region.bottom || current.right !== region.right)
        ? region : { top: 0, left: 0, bottom: limits.rows - 1, right: limits.cols - 1 };
      return moveAction({ ...selectionAt(state.focus), range: { ...range } });
    }
    return moveAction({
      anchor: { row: 0, col: 0 },
      focus: { row: limits.rows - 1, col: limits.cols - 1 },
    });
  }

  if (isPrintable(key)) {
    return { type: 'startEdit', initialInput: key.key };
  }

  return { type: 'none' };
}

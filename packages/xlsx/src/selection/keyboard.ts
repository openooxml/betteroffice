/**
 * Keyboard reducer: maps a keystroke over the current selection to a
 * declarative {@link SelectionAction}. Pure — it neither mutates state nor
 * touches the DOM; the chrome owns the keydown listener and interprets the
 * action. Excel navigation semantics, with two deliberate choices documented at
 * the F2/Enter handlers below.
 */

import { extendTo, moveFocus, selectionAt } from './model';
import type {
  CellAddr,
  Direction,
  KeyInput,
  Selection,
  SelectionAction,
  SelectionLimits,
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

// jump to the used-range edge in a direction (ctrl/cmd+arrow) — we have no data
// map here, so the boundary is the used range itself.
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

// tab/shift-tab: step within the row, wrapping to the next/previous row's far
// edge at the used-range bounds. always collapses (excel commits the cell).
function tabMove(sel: Selection, back: boolean, limits: SelectionLimits): Selection {
  const { row, col } = sel.focus;
  if (!back && col >= limits.cols - 1) {
    return selectionAt({ row: Math.min(row + 1, limits.rows - 1), col: 0 });
  }
  if (back && col <= 0) {
    return selectionAt({ row: Math.max(row - 1, 0), col: limits.cols - 1 });
  }
  return selectionAt({ row, col: col + (back ? -1 : 1) });
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
  limits: SelectionLimits
): SelectionAction {
  const direction = ARROWS[key.key];
  if (direction) {
    if (mod(key)) {
      const edge = edgeInDirection(state.focus, direction, limits);
      return moveAction(key.shiftKey ? extendTo(state, edge, limits) : selectionAt(edge));
    }
    return moveAction(moveFocus(state, direction, { extend: key.shiftKey, limits }));
  }

  switch (key.key) {
    case 'Tab':
      return moveAction(tabMove(state, Boolean(key.shiftKey), limits));

    // enter navigates (down/up); it never opens the editor — F2 does. this is
    // the excel split: type to overwrite, F2 to edit in place, enter to commit.
    case 'Enter':
      return moveAction(moveFocus(state, key.shiftKey ? 'up' : 'down', { extend: false, limits }));

    case 'Home':
      return moveAction(
        mod(key) ? selectionAt({ row: 0, col: 0 }) : selectionAt({ row: state.focus.row, col: 0 })
      );

    case 'End':
      return moveAction(
        mod(key)
          ? selectionAt({ row: limits.rows - 1, col: limits.cols - 1 })
          : selectionAt({ row: state.focus.row, col: limits.cols - 1 })
      );

    case 'PageDown':
      return moveAction(
        moveFocus(state, 'down', { extend: key.shiftKey, step: limits.rowsPerPage, limits })
      );

    case 'PageUp':
      return moveAction(
        moveFocus(state, 'up', { extend: key.shiftKey, step: limits.rowsPerPage, limits })
      );

    // F2 opens the editor on the current cell without seeding input.
    case 'F2':
      return { type: 'startEdit' };

    case 'Delete':
    case 'Backspace':
      return { type: 'clear' };
  }

  // cmd/ctrl+A selects the whole used range.
  if ((key.key === 'a' || key.key === 'A') && mod(key)) {
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

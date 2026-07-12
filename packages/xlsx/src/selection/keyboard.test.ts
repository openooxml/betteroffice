import { describe, expect, it } from 'bun:test';
import { selectionAt, selectionKeyReducer } from './index';
import type { KeyInput, Selection, SelectionAction, SelectionLimits } from './index';

const limits: SelectionLimits = { rows: 100, cols: 26, rowsPerPage: 10 };

function press(state: Selection, key: Partial<KeyInput> & { key: string }): SelectionAction {
  return selectionKeyReducer(state, key, limits);
}

// pull the selection out of a move action, failing loudly on any other action.
function moved(action: SelectionAction): Selection {
  if (action.type !== 'move') throw new Error(`expected move, got ${action.type}`);
  return action.selection;
}

describe('arrow navigation', () => {
  const sel = selectionAt({ row: 5, col: 5 });

  it('moves the collapsed selection with a bare arrow', () => {
    expect(moved(press(sel, { key: 'ArrowDown' }))).toEqual(selectionAt({ row: 6, col: 5 }));
    expect(moved(press(sel, { key: 'ArrowRight' }))).toEqual(selectionAt({ row: 5, col: 6 }));
  });

  it('extends with shift+arrow, holding the anchor', () => {
    const out = moved(press(sel, { key: 'ArrowDown', shiftKey: true }));
    expect(out).toEqual({ anchor: { row: 5, col: 5 }, focus: { row: 6, col: 5 } });
  });

  it('jumps to the used-range edge with ctrl/cmd+arrow', () => {
    expect(moved(press(sel, { key: 'ArrowDown', metaKey: true }))).toEqual(
      selectionAt({ row: 99, col: 5 })
    );
    expect(moved(press(sel, { key: 'ArrowLeft', ctrlKey: true }))).toEqual(
      selectionAt({ row: 5, col: 0 })
    );
  });

  it('extends to the edge with shift+ctrl+arrow', () => {
    const out = moved(press(sel, { key: 'ArrowRight', metaKey: true, shiftKey: true }));
    expect(out).toEqual({ anchor: { row: 5, col: 5 }, focus: { row: 5, col: 25 } });
  });
});

describe('tab and enter', () => {
  it('tab moves right and wraps to the next row start at the edge', () => {
    expect(moved(press(selectionAt({ row: 2, col: 3 }), { key: 'Tab' }))).toEqual(
      selectionAt({ row: 2, col: 4 })
    );
    expect(moved(press(selectionAt({ row: 2, col: 25 }), { key: 'Tab' }))).toEqual(
      selectionAt({ row: 3, col: 0 })
    );
  });

  it('shift+tab moves left and wraps to the previous row end', () => {
    expect(moved(press(selectionAt({ row: 2, col: 0 }), { key: 'Tab', shiftKey: true }))).toEqual(
      selectionAt({ row: 1, col: 25 })
    );
  });

  it('enter moves down and shift+enter moves up, never editing', () => {
    expect(moved(press(selectionAt({ row: 2, col: 3 }), { key: 'Enter' }))).toEqual(
      selectionAt({ row: 3, col: 3 })
    );
    expect(moved(press(selectionAt({ row: 2, col: 3 }), { key: 'Enter', shiftKey: true }))).toEqual(
      selectionAt({ row: 1, col: 3 })
    );
  });
});

describe('home/end and paging', () => {
  const sel = selectionAt({ row: 40, col: 10 });

  it('home goes to the row start, ctrl+home to the sheet origin', () => {
    expect(moved(press(sel, { key: 'Home' }))).toEqual(selectionAt({ row: 40, col: 0 }));
    expect(moved(press(sel, { key: 'Home', metaKey: true }))).toEqual(
      selectionAt({ row: 0, col: 0 })
    );
  });

  it('end goes to the row end, ctrl+end to the last used cell', () => {
    expect(moved(press(sel, { key: 'End' }))).toEqual(selectionAt({ row: 40, col: 25 }));
    expect(moved(press(sel, { key: 'End', ctrlKey: true }))).toEqual(
      selectionAt({ row: 99, col: 25 })
    );
  });

  it('pageDown/pageUp step by rowsPerPage', () => {
    expect(moved(press(sel, { key: 'PageDown' }))).toEqual(selectionAt({ row: 50, col: 10 }));
    expect(moved(press(sel, { key: 'PageUp' }))).toEqual(selectionAt({ row: 30, col: 10 }));
  });

  it('pageDown extends when shifted', () => {
    const out = moved(press(sel, { key: 'PageDown', shiftKey: true }));
    expect(out).toEqual({ anchor: { row: 40, col: 10 }, focus: { row: 50, col: 10 } });
  });
});

describe('select-all, edit, and clear', () => {
  const sel = selectionAt({ row: 3, col: 3 });

  it('cmd/ctrl+A selects the whole used range', () => {
    expect(moved(press(sel, { key: 'a', metaKey: true }))).toEqual({
      anchor: { row: 0, col: 0 },
      focus: { row: 99, col: 25 },
    });
    expect(moved(press(sel, { key: 'A', ctrlKey: true }))).toEqual({
      anchor: { row: 0, col: 0 },
      focus: { row: 99, col: 25 },
    });
  });

  it('F2 starts editing without seeding input', () => {
    expect(press(sel, { key: 'F2' })).toEqual({ type: 'startEdit' });
  });

  it('a printable key starts editing seeded with the character', () => {
    expect(press(sel, { key: 'x' })).toEqual({ type: 'startEdit', initialInput: 'x' });
    expect(press(sel, { key: '5' })).toEqual({ type: 'startEdit', initialInput: '5' });
  });

  it('does not treat a modified letter as printable input', () => {
    // a bare 'a' is printable; cmd+a is select-all, handled above, not startEdit.
    expect(press(sel, { key: 'a' })).toEqual({ type: 'startEdit', initialInput: 'a' });
    expect(press(sel, { key: 'c', metaKey: true }).type).toBe('none');
  });

  it('delete and backspace clear the selection contents', () => {
    expect(press(sel, { key: 'Delete' })).toEqual({ type: 'clear' });
    expect(press(sel, { key: 'Backspace' })).toEqual({ type: 'clear' });
  });

  it('returns none for keys it does not own', () => {
    expect(press(sel, { key: 'Escape' }).type).toBe('none');
    expect(press(sel, { key: 'F5' }).type).toBe('none');
  });
});

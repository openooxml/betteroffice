import { describe, expect, it } from 'bun:test';
import { cellCount, normalizeRange, selectionAt, selectionKeyReducer } from './index';
import type { KeyInput, Selection, SelectionAction, SelectionHidden, SelectionLimits } from './index';

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
    expect(moved(selectionKeyReducer(sel, { key: 'End', ctrlKey: true }, limits,
      { usedRange: { top: 0, left: 0, bottom: 99, right: 25 } }))).toEqual(
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

describe('data-aware keyboard selection', () => {
  const active = { row: 2, col: 2 };
  const state = selectionAt(active);
  const region = { top: 0, left: 0, bottom: 5, right: 3 };

  for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
    it(`uses the data edge in every direction with ${Object.keys(modifier)[0]}`, () => {
      for (const [key, edge] of [
        ['ArrowRight', { row: 2, col: 3 }],
        ['ArrowLeft', { row: 2, col: 0 }],
        ['ArrowDown', { row: 5, col: 2 }],
        ['ArrowUp', { row: 0, col: 2 }],
      ] as const) {
        const data = { edge };
        expect(moved(selectionKeyReducer(state, { key, ...modifier }, limits, data)))
          .toEqual(selectionAt(edge));
        expect(moved(selectionKeyReducer(state, { key, ...modifier, shiftKey: true }, limits, data)))
          .toEqual({ anchor: active, focus: edge });
      }
    });

    it('uses a next occupied cell beyond a gap or outside the viewport', () => {
      for (const edge of [{ row: 2, col: 7 }, { row: 90, col: 2 }, { row: 400, col: 2 }, { row: 2, col: 60 }]) {
        const key = edge.row === 2 ? 'ArrowRight' : 'ArrowDown';
        expect(moved(selectionKeyReducer(state, { key, ...modifier }, limits, { edge })))
          .toEqual(selectionAt(edge));
        expect(moved(selectionKeyReducer(state, { key, ...modifier, shiftKey: true }, limits, { edge })))
          .toEqual({ anchor: active, focus: edge });
      }
    });

    it('keeps the original anchor when extending an existing selection', () => {
      const selection = { anchor: { row: 1, col: 1 }, focus: active };
      expect(moved(selectionKeyReducer(selection, { key: 'ArrowDown', ...modifier, shiftKey: true },
        limits, { edge: { row: 5, col: 2 } })))
        .toEqual({ anchor: selection.anchor, focus: { row: 5, col: 2 } });
    });

    it('falls back to each existing grid edge when no data is found', () => {
      for (const [key, edge] of [
        ['ArrowRight', { row: 2, col: 25 }], ['ArrowLeft', { row: 2, col: 0 }],
        ['ArrowDown', { row: 99, col: 2 }], ['ArrowUp', { row: 0, col: 2 }],
      ] as const) {
        expect(moved(selectionKeyReducer(state, { key, ...modifier }, limits, { edge: null })))
          .toEqual(selectionAt(edge));
      }
    });

    it('selects the current region first and the full extent next without moving the active cell', () => {
      const key = { key: 'a', ...modifier };
      const first = moved(selectionKeyReducer(state, key, limits, { region }));
      expect(normalizeRange(first)).toEqual(region);
      expect(first.focus).toEqual(active);
      expect(cellCount(first)).toBe(24);
      const second = moved(selectionKeyReducer(first, key, limits, { region }));
      expect(normalizeRange(second)).toEqual({ top: 0, left: 0, bottom: 99, right: 25 });
      expect(second.focus).toEqual(active);
      expect(first).toEqual({ ...state, range: region });
    });

    it('recognizes an already selected region in either anchor direction', () => {
      for (const selection of [
        { anchor: { row: 0, col: 0 }, focus: { row: 5, col: 3 } },
        { anchor: { row: 5, col: 3 }, focus: { row: 0, col: 0 } },
      ]) {
        const next = moved(selectionKeyReducer(selection, { key: 'A', ...modifier }, limits, { region }));
        expect(normalizeRange(next)).toEqual({ top: 0, left: 0, bottom: 99, right: 25 });
        expect(next.focus).toEqual(selection.focus);
      }
    });

    it('selects everything when there is no current region while preserving focus', () => {
      const next = moved(selectionKeyReducer(state, { key: 'a', ...modifier }, limits, { region: null }));
      expect(normalizeRange(next)).toEqual({ top: 0, left: 0, bottom: 99, right: 25 });
      expect(next.focus).toEqual(active);
    });
  }

  it('resumes navigation from the active cell and drops explicit region bounds', () => {
    const selected = moved(selectionKeyReducer(state, { key: 'a', ctrlKey: true }, limits, { region }));
    expect(moved(selectionKeyReducer(selected, { key: 'ArrowRight' }, limits)))
      .toEqual(selectionAt({ row: 2, col: 3 }));
    expect(moved(selectionKeyReducer(selected, { key: 'ArrowDown', shiftKey: true }, limits)))
      .toEqual({ anchor: active, focus: { row: 3, col: 2 } });
  });
});

describe('hidden-track navigation', () => {
  const bounds = { rows: 6, cols: 6, rowsPerPage: 2 };
  const hidden: SelectionHidden = {
    isRowHidden: (row) => row === 1 || row === 3,
    isColHidden: (col) => col === 1 || col === 3,
  };
  const origin = selectionAt({ row: 0, col: 0 });
  const reduce = (state: Selection, key: KeyInput, tracks = hidden) =>
    moved(selectionKeyReducer(state, key, bounds, undefined, tracks));

  for (const [key, reverse, focus] of [
    ['ArrowRight', { key: 'ArrowLeft' }, { row: 0, col: 2 }],
    ['ArrowDown', { key: 'ArrowUp' }, { row: 2, col: 0 }],
    ['Tab', { key: 'Tab', shiftKey: true }, { row: 0, col: 2 }],
    ['Enter', { key: 'Enter', shiftKey: true }, { row: 2, col: 0 }],
  ] as const) {
    it(`${key} and its reverse skip hidden tracks`, () => {
      const next = reduce(origin, { key });
      expect(next).toEqual(selectionAt(focus));
      expect(reduce(next, reverse)).toEqual(origin);
    });
  }

  for (const [key, focus] of [
    ['ArrowRight', { row: 0, col: 2 }], ['ArrowDown', { row: 2, col: 0 }],
  ] as const) {
    it(`Shift+${key} extends across hidden tracks`, () => {
      expect(reduce(origin, { key, shiftKey: true })).toEqual({ anchor: origin.anchor, focus });
      const reverse = key === 'ArrowRight' ? 'ArrowLeft' : 'ArrowUp';
      expect(reduce(selectionAt(focus), { key: reverse, shiftKey: true }))
        .toEqual({ anchor: focus, focus: origin.focus });
    });
  }

  it('counts visible rows when paging and preserves the extension anchor', () => {
    const next = reduce(origin, { key: 'PageDown', shiftKey: true });
    expect(next).toEqual({ anchor: origin.anchor, focus: { row: 4, col: 0 } });
    expect(reduce(next, { key: 'PageUp' })).toEqual(origin);
  });

  it('wraps Tab across hidden rows and leading or trailing columns', () => {
    const forward = { isRowHidden: hidden.isRowHidden, isColHidden: (col: number) => col === 0 };
    const backward = { isRowHidden: hidden.isRowHidden, isColHidden: (col: number) => col === 5 };
    expect(reduce(selectionAt({ row: 0, col: 5 }), { key: 'Tab' }, forward))
      .toEqual(selectionAt({ row: 2, col: 1 }));
    expect(reduce(selectionAt({ row: 2, col: 0 }), { key: 'Tab', shiftKey: true }, backward))
      .toEqual(selectionAt({ row: 0, col: 4 }));
  });

  it('wraps Tab when every column after the active cell is hidden', () => {
    const tail = { isRowHidden: hidden.isRowHidden, isColHidden: (col: number) => col >= 4 };
    expect(reduce(selectionAt({ row: 0, col: 3 }), { key: 'Tab' }, tail))
      .toEqual(selectionAt({ row: 2, col: 0 }));
    expect(reduce(selectionAt({ row: 2, col: 0 }), { key: 'Tab', shiftKey: true }, tail))
      .toEqual(selectionAt({ row: 0, col: 3 }));
  });

  it('uses visible sheet and row corners for Home and End', () => {
    const tracks = { isRowHidden: (row: number) => row === 0 || row === 5,
      isColHidden: (col: number) => col === 0 || col === 5 };
    const state = selectionAt({ row: 2, col: 2 });
    expect(reduce(state, { key: 'Home' }, tracks)).toEqual(selectionAt({ row: 2, col: 1 }));
    expect(reduce(state, { key: 'Home', ctrlKey: true }, tracks)).toEqual(selectionAt({ row: 1, col: 1 }));
    expect(reduce(state, { key: 'End' }, tracks)).toEqual(selectionAt({ row: 2, col: 4 }));
  });

  for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
    it(`skips hidden data targets and trims hidden fallback edges with ${Object.keys(modifier)[0]}`, () => {
      expect(moved(selectionKeyReducer(origin, { key: 'ArrowRight', ...modifier }, bounds,
        { edge: { row: 0, col: 1 } }, hidden))).toEqual(selectionAt({ row: 0, col: 2 }));
      const tracks = { isRowHidden: (row: number) => row === 0 || row >= 4,
        isColHidden: (col: number) => col === 0 || col >= 4 };
      const state = selectionAt({ row: 2, col: 2 });
      for (const [key, focus] of [
        ['ArrowRight', { row: 2, col: 3 }], ['ArrowDown', { row: 3, col: 2 }],
        ['ArrowLeft', { row: 2, col: 1 }], ['ArrowUp', { row: 1, col: 2 }],
      ] as const) {
        expect(moved(selectionKeyReducer(state, { key, ...modifier, shiftKey: true }, bounds,
          { edge: null }, tracks))).toEqual({ anchor: state.anchor, focus });
      }
    });
  }

  it('preserves the selection when every track in the direction is hidden', () => {
    const state = { ...origin, range: { top: 0, left: 0, bottom: 2, right: 2 } };
    const tracks = { isRowHidden: (row: number) => row > 0, isColHidden: (col: number) => col > 0 };
    for (const key of ['ArrowRight', 'ArrowDown', 'Enter', 'Tab', 'PageDown']) {
      expect(reduce(state, { key }, tracks)).toEqual(state);
      expect(reduce(state, { key, shiftKey: true }, tracks)).toEqual(state);
    }
    for (const key of ['ArrowRight', 'ArrowDown']) {
      expect(reduce(state, { key, ctrlKey: true, shiftKey: true }, tracks)).toEqual(state);
    }
  });
});

describe('sheet bounds and used range', () => {
  const bounds = { rows: 1_048_576, cols: 16_384, rowsPerPage: 20 };
  const state = selectionAt({ row: 9_999, col: 701 });
  const usedRange = { top: 0, left: 0, bottom: 9_999, right: 701 };

  for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
    it(`uses the used rectangle for ${Object.keys(modifier)[0]}+End independently of the current region`, () => {
      const start = selectionAt({ row: 50_000, col: 2_000 });
      const data = { usedRange, region: { top: 0, left: 0, bottom: 0, right: 0 } };
      expect(moved(selectionKeyReducer(start, { key: 'End', ...modifier }, bounds, data))).toEqual(state);
      expect(moved(selectionKeyReducer(start, { key: 'End', ...modifier, shiftKey: true }, bounds, data)))
        .toEqual({ anchor: start.anchor, focus: state.focus });
      expect(moved(selectionKeyReducer(start, { key: 'Home', ...modifier }, bounds, data)))
        .toEqual(selectionAt({ row: 0, col: 0 }));
    });

    it('uses A1 for an empty used range and ignores unavailable data', () => {
      expect(moved(selectionKeyReducer(state, { key: 'End', ...modifier }, bounds, { usedRange: null })))
        .toEqual(selectionAt({ row: 0, col: 0 }));
      for (const data of [undefined, {}, { region: null }]) {
        expect(selectionKeyReducer(state, { key: 'End', ...modifier }, bounds, data)).toEqual({ type: 'none' });
      }
    });

    it('continues from the last data cell to the real sheet edge', () => {
      for (const [key, focus] of [
        ['ArrowDown', { row: 1_048_575, col: 701 }],
        ['ArrowRight', { row: 9_999, col: 16_383 }],
      ] as const) {
        expect(moved(selectionKeyReducer(state, { key, ...modifier }, bounds, { edge: null, usedRange })))
          .toEqual(selectionAt(focus));
      }
    });
  }

  it('allows arrows, Shift arrows, paging and End beyond the used range', () => {
    for (const [key, focus] of [
      ['ArrowDown', { row: 10_000, col: 701 }],
      ['ArrowRight', { row: 9_999, col: 702 }],
      ['PageDown', { row: 10_019, col: 701 }],
      ['End', { row: 9_999, col: 16_383 }],
    ] as const) {
      expect(moved(selectionKeyReducer(state, { key }, bounds))).toEqual(selectionAt(focus));
      if (key !== 'End') expect(moved(selectionKeyReducer(state, { key, shiftKey: true }, bounds)))
        .toEqual({ anchor: state.anchor, focus });
    }
    const last = selectionAt({ row: bounds.rows - 1, col: bounds.cols - 1 });
    for (const key of ['ArrowDown', 'ArrowRight', 'PageDown', 'End']) {
      expect(moved(selectionKeyReducer(last, { key }, bounds))).toEqual(last);
    }
  });

  it('keeps a hidden used-range corner and clamps malformed range bounds to the sheet', () => {
    const hidden = { isRowHidden: () => true, isColHidden: () => true };
    expect(moved(selectionKeyReducer(selectionAt({ row: 0, col: 0 }), { key: 'End', ctrlKey: true },
      bounds, { usedRange }, hidden))).toEqual(state);
    expect(moved(selectionKeyReducer(state, { key: 'End', ctrlKey: true }, bounds,
      { usedRange: { ...usedRange, bottom: bounds.rows, right: bounds.cols } })))
      .toEqual(selectionAt({ row: bounds.rows - 1, col: bounds.cols - 1 }));
  });
});

import { describe, expect, it } from 'bun:test';
import { cellCount, normalizeRange, selectionAt, selectionKeyReducer } from './index';
import type { CellRange, KeyInput, Selection, SelectionAction, SelectionData, SelectionHidden, SelectionLimits, SelectionMerges } from './index';

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

  it('end goes to the last data cell in the row, ctrl+end to the last used cell', () => {
    const data = { edge: { row: 40, col: 17 } };
    expect(moved(selectionKeyReducer(sel, { key: 'End' }, limits, data))).toEqual(selectionAt({ row: 40, col: 17 }));
    expect(moved(selectionKeyReducer(sel, { key: 'End', shiftKey: true }, limits, data)))
      .toEqual({ anchor: sel.anchor, focus: { row: 40, col: 17 } });
    expect(selectionKeyReducer(sel, { key: 'End', ctrlKey: true }, limits, data)).toEqual({ type: 'none' });
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

  it('uses visible sheet and row corners for Home and the visible data target for End', () => {
    const tracks = { isRowHidden: (row: number) => row === 0 || row === 5,
      isColHidden: (col: number) => col === 0 || col === 5 };
    const state = selectionAt({ row: 2, col: 2 });
    expect(reduce(state, { key: 'Home' }, tracks)).toEqual(selectionAt({ row: 2, col: 1 }));
    expect(reduce(state, { key: 'Home', ctrlKey: true }, tracks)).toEqual(selectionAt({ row: 1, col: 1 }));
    expect(moved(selectionKeyReducer(state, { key: 'End' }, bounds, { edge: { row: 2, col: 4 } }, tracks)))
      .toEqual(selectionAt({ row: 2, col: 4 }));
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

  it('allows arrows, Shift arrows and paging beyond the used range', () => {
    for (const [key, focus] of [
      ['ArrowDown', { row: 10_000, col: 701 }],
      ['ArrowRight', { row: 9_999, col: 702 }],
      ['PageDown', { row: 10_019, col: 701 }],
    ] as const) {
      expect(moved(selectionKeyReducer(state, { key }, bounds))).toEqual(selectionAt(focus));
      expect(moved(selectionKeyReducer(state, { key, shiftKey: true }, bounds)))
        .toEqual({ anchor: state.anchor, focus });
    }
    const last = selectionAt({ row: bounds.rows - 1, col: bounds.cols - 1 });
    for (const key of ['ArrowDown', 'ArrowRight', 'PageDown']) {
      expect(moved(selectionKeyReducer(last, { key }, bounds))).toEqual(last);
    }
  });

  it('keeps plain End inside the data: stays put without a target and ignores unavailable data', () => {
    const range = { anchor: { row: 9_999, col: 3 }, focus: state.focus };
    for (const shiftKey of [false, true]) {
      expect(moved(selectionKeyReducer(state, { key: 'End', shiftKey }, bounds, { edge: null }))).toBe(state);
      expect(moved(selectionKeyReducer(range, { key: 'End', shiftKey }, bounds, { edge: null }))).toBe(range);
      for (const data of [undefined, {}, { usedRange }]) {
        expect(selectionKeyReducer(state, { key: 'End', shiftKey }, bounds, data)).toEqual({ type: 'none' });
      }
    }
    expect(moved(selectionKeyReducer(state, { key: 'End' }, bounds, { edge: { row: 9_999, col: bounds.cols } })))
      .toEqual(selectionAt({ row: 9_999, col: bounds.cols - 1 }));
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

describe('merged-area navigation', () => {
  const areas: CellRange[] = [
    { top: 2, left: 2, bottom: 3, right: 4 },
    { top: 6, left: 2, bottom: 7, right: 3 },
    { top: 0, left: 24, bottom: 1, right: 25 },
    { top: 10, left: 0, bottom: 11, right: 1 },
    { top: 98, left: 0, bottom: 99, right: 0 },
  ];
  const merges: SelectionMerges = {
    mergeAt: (row, col) => areas.find((area) => row >= area.top && row <= area.bottom &&
      col >= area.left && col <= area.right) ?? null,
  };
  const at = (row: number, col: number) => selectionAt({ row, col });
  const reduce = (state: Selection, key: Partial<KeyInput> & { key: string }, data?: SelectionData,
    hidden?: SelectionHidden) => moved(selectionKeyReducer(state, key, limits, data, hidden, merges));

  it.each([
    ['ArrowDown', at(2, 2), at(4, 2)],
    ['ArrowUp', at(4, 2), at(2, 2)],
    ['ArrowUp', at(2, 2), at(1, 2)],
    ['ArrowDown', at(1, 3), at(2, 2)],
    ['ArrowRight', at(2, 2), at(2, 5)],
    ['ArrowLeft', at(2, 5), at(2, 2)],
    ['ArrowLeft', at(3, 5), at(2, 2)],
    ['ArrowLeft', at(2, 2), at(2, 1)],
    ['ArrowRight', at(3, 1), at(2, 2)],
  ] as const)('%s from %o skips the cells a merge covers', (key, from, to) => {
    expect(reduce(from, { key })).toEqual(to);
  });

  it('Enter and Shift+Enter step over merges and land on their anchors', () => {
    const path = [at(0, 2)];
    for (let i = 0; i < 6; i++) path.push(reduce(path[path.length - 1]!, { key: 'Enter' }));
    expect(path.map((sel) => sel.focus.row)).toEqual([0, 1, 2, 4, 5, 6, 8]);
    expect(reduce(at(8, 2), { key: 'Enter', shiftKey: true })).toEqual(at(6, 2));
    expect(reduce(at(6, 2), { key: 'Enter', shiftKey: true })).toEqual(at(5, 2));
    expect(reduce(at(4, 2), { key: 'Enter', shiftKey: true })).toEqual(at(2, 2));
  });

  it('Tab and Shift+Tab step over merges and land on their anchors', () => {
    expect(reduce(at(2, 2), { key: 'Tab' })).toEqual(at(2, 5));
    expect(reduce(at(3, 1), { key: 'Tab' })).toEqual(at(2, 2));
    expect(reduce(at(2, 5), { key: 'Tab', shiftKey: true })).toEqual(at(2, 2));
    expect(reduce(at(2, 2), { key: 'Tab', shiftKey: true })).toEqual(at(2, 1));
  });

  it('wraps Tab past a merge on the sheet edge and lands on a merge anchor after wrapping', () => {
    expect(reduce(at(0, 24), { key: 'Tab' })).toEqual(at(2, 0));
    expect(reduce(at(10, 0), { key: 'Tab', shiftKey: true })).toEqual(at(9, 25));
    expect(reduce(at(10, 25), { key: 'Tab' })).toEqual(at(10, 0));
    expect(reduce(at(12, 0), { key: 'Tab', shiftKey: true })).toEqual(at(11, 25));
  });

  it('keeps a merge on the sheet edge selected and collapses a covered focus onto its anchor', () => {
    expect(reduce(at(98, 0), { key: 'ArrowDown' })).toEqual(at(98, 0));
    expect(reduce(at(99, 0), { key: 'Enter' })).toEqual(at(98, 0));
    expect(reduce(at(0, 24), { key: 'ArrowRight' })).toEqual(at(0, 24));
  });

  it('extends from the far edge of a merge', () => {
    expect(reduce(at(2, 2), { key: 'ArrowDown', shiftKey: true })).toEqual({ anchor: { row: 2, col: 2 }, focus: { row: 4, col: 2 } });
    expect(reduce(at(2, 2), { key: 'ArrowRight', shiftKey: true })).toEqual({ anchor: { row: 2, col: 2 }, focus: { row: 2, col: 5 } });
  });

  it('pages from the far edge of a merge', () => {
    expect(reduce(at(2, 2), { key: 'PageDown' })).toEqual(at(13, 2));
    expect(reduce(at(13, 3), { key: 'PageUp' })).toEqual(at(2, 2));
  });

  it('lands jumps, Home and End on merge anchors', () => {
    expect(reduce(at(3, 8), { key: 'ArrowLeft', ctrlKey: true }, { edge: { row: 3, col: 4 } })).toEqual(at(2, 2));
    expect(reduce(at(11, 5), { key: 'Home' })).toEqual(at(10, 0));
    expect(reduce(at(0, 0), { key: 'End', ctrlKey: true }, { usedRange: { top: 0, left: 0, bottom: 7, right: 3 } })).toEqual(at(6, 2));
  });

  it('skips hidden tracks past a merge', () => {
    const hidden: SelectionHidden = { isRowHidden: (row) => row === 4, isColHidden: (col) => col === 5 };
    expect(reduce(at(2, 2), { key: 'ArrowDown' }, undefined, hidden)).toEqual(at(5, 2));
    expect(reduce(at(2, 2), { key: 'ArrowRight' }, undefined, hidden)).toEqual(at(2, 6));
    expect(reduce(at(5, 2), { key: 'ArrowUp' }, undefined, hidden)).toEqual(at(2, 2));
  });
});

it('uses the workbook merge anchor for Ctrl+End without painted merges', () => {
  const state = selectionAt({ row: 0, col: 0 });
  const action = selectionKeyReducer(state, { key: 'End', ctrlKey: true },
    { rows: 1_048_576, cols: 16_384, rowsPerPage: 20 }, {
      usedRange: { top: 0, left: 0, bottom: 100, right: 4 }, edge: { row: 99, col: 3 },
    });
  expect(action).toEqual({ type: 'move', selection: selectionAt({ row: 99, col: 3 }) });
});

import { expect } from 'bun:test';

import { SHEET, fingerprint, profiledOps } from './context';
import type { XlsxScenario } from './context';

const TOP = 500;
const ROWS = 20;
const ANCHOR = 'STORM-ANCHOR';

export const structuralStorm: XlsxScenario = {
  name: 'structural-storm',
  description:
    'Twenty interleaved row inserts and deletes around a formula-bearing block, tracked through a moving anchor, then the whole storm undone and redone against fingerprints of every cell input.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const a1 = (row: number, col: number) => handle.cell(SHEET, row, col).a1;
    const seed = [{ row: TOP, col: 3, input: ANCHOR }];
    for (let row = 0; row < ROWS; row += 1) {
      seed.push({ row: TOP + row, col: 0, input: String(row + 1) });
      seed.push({ row: TOP + row, col: 1, input: `=${a1(TOP + row, 0)}*2` });
      seed.push({
        row: TOP + row,
        col: 2,
        input: `=SUM(${a1(TOP, 0)}:${a1(TOP + row, 0)})`,
      });
    }
    expect(
      recorder.op('editCells:seedRegion', () => handle.editCells(SHEET, seed))
        .applied
    ).toBe(true);

    const anchorRow = () => {
      const found = recorder.op('searchText:anchor', () =>
        handle.searchText(ANCHOR)
      );
      expect(found.length).toBe(1);
      return found[0].row;
    };
    const region = (top: number) => `${a1(top, 0)}:${a1(top + ROWS + 6, 3)}`;
    const before = fingerprint(handle, region(TOP));
    expect(before).toContain(ANCHOR);

    const round = (top: number) => [
      { type: 'insertRows', sheet: SHEET, at: top + 2, count: 2 },
      { type: 'deleteRows', sheet: SHEET, at: top + 6, count: 1 },
      { type: 'insertRows', sheet: SHEET, at: 0, count: 1 },
      { type: 'deleteRows', sheet: SHEET, at: 0, count: 1 },
      { type: 'insertRows', sheet: SHEET, at: top + ROWS, count: 1 },
    ];
    let applied = 0;
    for (let index = 0; index < 4; index += 1) {
      for (const op of round(anchorRow())) {
        expect(
          profiledOps(handle, recorder, `applyOps:${String(op.type)}`, [op])
            .applied
        ).toBe(true);
        applied += 1;
      }
    }
    expect(applied).toBe(20);
    expect(anchorRow()).toBe(TOP);

    const after = fingerprint(handle, region(TOP));
    expect(after).not.toBe(before);
    const formulas = handle
      .rangeCells(SHEET, region(TOP))
      .flat()
      .filter((cell) => cell.isFormula);
    expect(formulas.length).toBeGreaterThan(ROWS);
    expect(formulas.some((cell) => cell.input.includes('#REF!'))).toBe(false);

    expect(
      recorder.op('historyState:afterStorm', () => handle.historyState())
        .undoDepth
    ).toBeGreaterThanOrEqual(20);
    for (let step = 0; step < 20; step += 1) {
      expect(recorder.op('undo:storm', () => handle.undo()).applied).toBe(true);
    }
    expect(fingerprint(handle, region(TOP))).toBe(before);

    for (let step = 0; step < 20; step += 1) {
      expect(recorder.op('redo:storm', () => handle.redo()).applied).toBe(true);
    }
    expect(fingerprint(handle, region(TOP))).toBe(after);
    expect(
      recorder.op('historyState:afterRedo', () => handle.historyState())
        .redoDepth
    ).toBe(0);
  },
};

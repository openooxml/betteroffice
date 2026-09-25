import { expect } from 'bun:test';

import { SHEET, fingerprint, profiledEdit } from './context';
import type { XlsxScenario } from './context';

const ROW = 420;
const CYCLES = 10;

export const saveLoadCycles: XlsxScenario = {
  name: 'save-load-cycles',
  description:
    'Ten rounds of edit, save and reopen on the same workbook, checking cell-input persistence, sheet names, and bounded file growth while recording save/open timings.',
  participants: ['web'],
  run({ recorder, open }) {
    let handle = recorder.load(() => open());
    const a1 = (row: number, col: number) => handle.cell(SHEET, row, col).a1;
    const region = `${a1(ROW, 0)}:${a1(ROW + CYCLES, 2)}`;
    const names = handle.sheetInfo().sheetNames;
    let previousBytes = 0;

    for (let cycle = 0; cycle < CYCLES; cycle += 1) {
      expect(
        profiledEdit(
          handle,
          recorder,
          'editCell:cycle',
          ROW + cycle,
          0,
          `cycle ${cycle}`
        ).applied
      ).toBe(true);
      expect(
        profiledEdit(
          handle,
          recorder,
          'editCell:cycleFormula',
          ROW + cycle,
          1,
          `=LEN(${a1(ROW + cycle, 0)})`
        ).applied
      ).toBe(true);
      const expected = fingerprint(handle, region);

      const saved = recorder.op('save:cycle', () => handle.save());
      expect(saved.byteLength).toBeGreaterThan(0);
      if (previousBytes > 0)
        expect(Math.abs(saved.byteLength - previousBytes)).toBeLessThan(
          previousBytes * 0.25
        );
      previousBytes = saved.byteLength;

      handle = recorder.op('reopen:cycle', () => open(saved));
      expect(fingerprint(handle, region)).toBe(expected);
      expect(handle.sheetInfo().sheetNames).toEqual(names);
    }

    expect(handle.cell(SHEET, ROW + CYCLES - 1, 0).input).toBe(
      `cycle ${CYCLES - 1}`
    );
    expect(
      recorder.op('searchText:allCycles', () => handle.searchText('cycle '))
        .length
    ).toBeGreaterThanOrEqual(CYCLES);
  },
};

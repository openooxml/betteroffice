import { expect } from 'bun:test';

import { SHEET, displayedText, profiledEdit } from './context';
import type { XlsxScenario } from './context';

const ROW = 700;
const WORD = 'Quarterly revenue, EMEA region, provisional';

export const incrementalCommits: XlsxScenario = {
  name: 'incremental-commits',
  description:
    'Commits every prefix of a label and formula, exercising incremental parsing and recalculation. Browser draft typing is covered separately.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const seed = Array.from({ length: 10 }, (_, index) => ({
      row: ROW + 1 + index,
      col: 0,
      input: String(index + 1),
    }));
    expect(
      recorder.op('editCells:seedRange', () => handle.editCells(SHEET, seed))
        .applied
    ).toBe(true);
    const formula = `=SUM(${handle.cell(SHEET, ROW + 1, 0).a1}:${
      handle.cell(SHEET, ROW + 10, 0).a1
    })`;

    for (let length = 1; length <= WORD.length; length += 1) {
      const keystroke = profiledEdit(
        handle,
        recorder,
        'editCell:prefix',
        ROW,
        1,
        WORD.slice(0, length)
      );
      expect(keystroke.applied).toBe(true);
    }
    expect(handle.cell(SHEET, ROW, 1).input).toBe(WORD);

    for (let length = 1; length <= formula.length; length += 1) {
      const partial = formula.slice(0, length);
      const keystroke = profiledEdit(
        handle,
        recorder,
        'editCell:formulaPrefix',
        ROW,
        0,
        partial
      );
      expect(keystroke.applied).toBe(true);
      expect([partial, `'${partial}`]).toContain(
        handle.cell(SHEET, ROW, 0).input
      );
    }
    expect(handle.cell(SHEET, ROW, 0)).toMatchObject({
      input: formula,
      isFormula: true,
    });
    expect(handle.cell(SHEET, ROW, 0).input.startsWith("'")).toBe(false);
    expect(
      displayedText(handle, recorder, 'searchText:typedFormula', '55', ROW, 0)
    ).toContain('55');
  },
};

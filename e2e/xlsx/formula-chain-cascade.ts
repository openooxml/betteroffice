import { expect } from 'bun:test';

import { SHEET, displayedText, profiledEdit } from './context';
import type { XlsxScenario } from './context';

const TOP = 300;
const LENGTH = 200;
const SUM = TOP + LENGTH;
const MAX = SUM + 1;

export const formulaChainCascade: XlsxScenario = {
  name: 'formula-chain-cascade',
  description:
    'Builds a 200-link dependency chain with aggregates over it, then edits the root ten times and measures every cascade.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const a1 = (row: number) => handle.cell(SHEET, row, 0).a1;
    const chain = Array.from({ length: LENGTH }, (_, index) => ({
      row: TOP + index,
      col: 0,
      input: index === 0 ? '1' : `=${a1(TOP + index - 1)}+1`,
    }));
    const first = a1(TOP);
    const last = a1(TOP + LENGTH - 1);
    chain.push({ row: SUM, col: 0, input: `=SUM(${first}:${last})` });
    chain.push({ row: MAX, col: 0, input: `=MAX(${first}:${last})` });

    const built = recorder.op('editCells:buildChain', () =>
      handle.editCells(SHEET, chain)
    );
    expect(built.applied).toBe(true);
    expect(handle.cell(SHEET, TOP + LENGTH - 1, 0).input).toBe(
      `=${a1(TOP + LENGTH - 2)}+1`
    );
    expect(
      displayedText(
        handle,
        recorder,
        'searchText:chainTail',
        '200',
        TOP + LENGTH - 1,
        0
      )
    ).toContain('200');
    expect(
      displayedText(handle, recorder, 'searchText:sum', '20100', SUM, 0)
    ).toContain('20100');

    for (let step = 1; step <= 10; step += 1) {
      const root = profiledEdit(
        handle,
        recorder,
        'editCell:root',
        TOP,
        0,
        String(1 + step)
      );
      expect(root.applied).toBe(true);
      expect(root.changed!.length).toBe(LENGTH + 1);
      expect(root.changed).toContain(a1(SUM));
    }
    expect(
      displayedText(handle, recorder, 'searchText:cascadedSum', '22100', SUM, 0)
    ).toContain('22100');
    expect(
      displayedText(handle, recorder, 'searchText:cascadedMax', '210', MAX, 0)
    ).toContain('210');

    const undone = recorder.op('undo:lastRoot', () => handle.undo());
    expect(undone.applied).toBe(true);
    expect(handle.cell(SHEET, TOP, 0).input).toBe('10');
    expect(
      displayedText(
        handle,
        recorder,
        'searchText:rolledBackSum',
        '21900',
        SUM,
        0
      )
    ).toContain('21900');
  },
};

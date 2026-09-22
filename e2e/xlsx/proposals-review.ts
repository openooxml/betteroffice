import { expect } from 'bun:test';

import { isProposalsAvailable } from '../../packages/xlsx/src/wasm/loader';
import { SHEET, displayedText } from './context';
import type { XlsxScenario } from './context';

const ROW = 360;

export const proposalsReview: XlsxScenario = {
  name: 'proposals-review',
  description:
    'An agent proposes a batch of cell edits, the reviewer lists them with their before and after texts, accepts one and rejects the other, and the accepted value survives a save and reopen.',
  participants: ['web', 'agent'],
  requires: () =>
    isProposalsAvailable()
      ? undefined
      : 'the core was built without the proposals api',
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const a1 = (row: number, col: number) => handle.cell(SHEET, row, col).a1;
    expect(
      recorder.op('editCells:base', () =>
        handle.editCells(SHEET, [
          { row: ROW, col: 0, input: '100' },
          { row: ROW + 1, col: 0, input: '200' },
        ])
      ).applied
    ).toBe(true);

    const accepted = recorder.op(
      'propose:accepted',
      () =>
        handle.propose('agent-a', 'raise the first figure', [
          { sheet: SHEET, row: ROW, col: 0, input: '150' },
        ]),
      undefined,
      { actor: 'agent' }
    );
    const rejected = recorder.op(
      'propose:rejected',
      () =>
        handle.propose('agent-b', 'zero the second figure', [
          { sheet: SHEET, row: ROW + 1, col: 0, input: '0' },
        ]),
      undefined,
      { actor: 'agent' }
    );
    expect(accepted.cells[0]).toMatchObject({
      a1: a1(ROW, 0),
      oldText: '100',
      newText: '150',
    });
    expect(rejected.cells[0]).toMatchObject({
      a1: a1(ROW + 1, 0),
      oldText: '200',
      newText: '0',
    });

    const listed = recorder.op('listProposals', () => handle.listProposals());
    expect(listed.map((proposal) => proposal.id).sort()).toEqual(
      [accepted.id, rejected.id].sort()
    );
    expect(handle.cell(SHEET, ROW, 0).input).toBe('100');

    const applied = recorder.op('acceptProposal', () =>
      handle.acceptProposal(accepted.id)
    );
    expect(applied.applied).toBe(true);
    expect(handle.cell(SHEET, ROW, 0).input).toBe('150');
    expect(
      displayedText(handle, recorder, 'searchText:accepted', '150', ROW, 0)
    ).toContain('150');

    expect(
      recorder.op('rejectProposal', () => handle.rejectProposal(rejected.id))
    ).toBe(true);
    expect(
      recorder.op('listProposals:afterReview', () => handle.listProposals())
    ).toEqual([]);
    expect(handle.cell(SHEET, ROW + 1, 0).input).toBe('200');

    const saved = recorder.op('save', () => handle.save());
    const reopened = recorder.op('reopen', () => open(saved));
    expect(reopened.cell(SHEET, ROW, 0).input).toBe('150');
    expect(reopened.cell(SHEET, ROW + 1, 0).input).toBe('200');
  },
};

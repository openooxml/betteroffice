import { expect } from 'bun:test';

import { SHEET, editStages, exchange, fingerprint, replicas } from './context';
import type { XlsxScenario } from './context';

const ROW = 320;
const REGION = 'A321:F330';

export const twoEditorsConverge: XlsxScenario = {
  name: 'two-editors-converge',
  description:
    'Two collaborative replicas edit disjoint cells, then the same cell, then styles and number formats, and every update is delivered both ways until both sides hold identical inputs, formats and search results.',
  participants: ['web:a', 'web:b'],
  run(ctx) {
    const [a, b] = replicas(ctx, ['web:a', 'web:b']);
    const a1 = (row: number, col: number) => a.handle.cell(SHEET, row, col).a1;

    const profiled = (
      peer: typeof a,
      op: string,
      row: number,
      col: number,
      input: string
    ) => {
      let profile;
      const result = peer.timer.op(
        op,
        () => {
          const edit = peer.handle.editCellProfiled(SHEET, row, col, input);
          profile = edit.profile;
          return edit;
        },
        () => editStages(profile!)
      );
      expect(result.applied).toBe(true);
      return result;
    };

    profiled(a, 'editCell:disjoint', ROW, 0, 'from A');
    profiled(b, 'editCell:disjoint', ROW, 1, 'from B');
    profiled(a, 'editCell:formula', ROW + 1, 0, `=${a1(ROW, 0)}&" seen"`);
    const first = exchange([a, b]);
    expect(first).toBeGreaterThan(0);
    expect(fingerprint(a.handle, REGION)).toBe(fingerprint(b.handle, REGION));
    expect(b.handle.cell(SHEET, ROW, 0).input).toBe('from A');
    expect(a.handle.cell(SHEET, ROW, 1).input).toBe('from B');

    expect(
      a.timer.op('patchRangeStyle', () =>
        a.handle.patchRangeStyle(SHEET, `${a1(ROW, 0)}:${a1(ROW, 2)}`, {
          bold: true,
          fillColor: '#ffcc00',
        })
      ).applied
    ).toBe(true);
    expect(
      b.timer.op('setNumberFormat', () =>
        b.handle.setNumberFormat(
          SHEET,
          `${a1(ROW + 2, 0)}:${a1(ROW + 4, 0)}`,
          'percent'
        )
      ).applied
    ).toBe(true);
    exchange([a, b]);
    expect(
      b.timer.op('selectionFormatting', () =>
        b.handle.selectionFormatting(SHEET, a1(ROW, 0))
      ).bold
    ).toBe(true);
    expect(
      a.timer.op('selectionFormatting', () =>
        a.handle.selectionFormatting(SHEET, a1(ROW + 2, 0))
      ).numberFormat
    ).toBe('percent');

    profiled(a, 'editCell:conflict', ROW + 5, 0, 'A wins?');
    profiled(b, 'editCell:conflict', ROW + 5, 0, 'B wins?');
    exchange([a, b]);
    const resolved = a.handle.cell(SHEET, ROW + 5, 0).input;
    expect(['A wins?', 'B wins?']).toContain(resolved);
    expect(b.handle.cell(SHEET, ROW + 5, 0).input).toBe(resolved);
    expect(fingerprint(a.handle, REGION)).toBe(fingerprint(b.handle, REGION));

    const probe = a.timer.op('searchText:probe', () =>
      a.handle.searchText('seen')
    );
    expect(
      b.timer
        .op('searchText:probe', () => b.handle.searchText('seen'))
        .map((match) => match.a1)
    ).toEqual(probe.map((match) => match.a1));
    expect([...a.handle.encodeStateVector()].join()).toBe(
      [...b.handle.encodeStateVector()].join()
    );

    const catchUp = a.timer.op(
      'encodeStateAsUpdate:inSync',
      () => a.handle.encodeStateAsUpdate(b.handle.encodeStateVector()),
      undefined,
      {
        bytes: a.handle.encodeStateAsUpdate(b.handle.encodeStateVector())
          .byteLength,
      }
    );
    expect(catchUp.byteLength).toBeLessThan(
      a.handle.encodeStateAsUpdate().byteLength
    );
  },
};

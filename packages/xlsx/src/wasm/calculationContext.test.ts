import { beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { WorkbookCalculationContext, WorkbookHandle, XlsxEditRequest } from '../index';
import { initWasm, openWorkbook } from './loader';

const CONTEXT: WorkbookCalculationContext = { nowSerial: 45_000.75, randSeed: 42 };
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');
const VIEWPORT = { x: 0, y: 0, width: 800, height: 800 };

function localSerial(ms: number): number {
  return (ms - new Date(ms).getTimezoneOffset() * 60_000) / 86_400_000 + 25_569;
}

function bytes(name = 'sample'): Uint8Array {
  return new Uint8Array(readFileSync(resolve(import.meta.dir, `../../test-fixtures/${name}.xlsx`)));
}

function volatileBytes(name = 'sample'): Uint8Array {
  const handle = openWorkbook(bytes(name), { calculation: { nowSerial: 1, randSeed: 0 } });
  try {
    handle.editCells(0, [
      '=NOW()', '=TODAY()', '=RANDBETWEEN(1,1000000)', '=RANDBETWEEN(1,1000000)',
      '=RANDBETWEEN(1,1000000)+RANDBETWEEN(1,1000000)',
    ].map((input, col) => ({ row: 30, col, input })));
    return handle.save();
  } finally {
    handle.dispose();
  }
}

function values(handle: WorkbookHandle): unknown[] {
  const read = handle.readCells({
    ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A31:E31' } }],
  });
  if (!read.ok) throw new Error(read.failure.message);
  return read.ranges[0].cells[0].map((cell) => cell.value);
}

function request(handle: WorkbookHandle, input: string, calculation?: XlsxEditRequest['calculation']): XlsxEditRequest {
  return {
    expectVersion: handle.version(),
    ...(calculation === undefined ? {} : { calculation }),
    steps: [{
      op: 'setCellInputs',
      target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'Z40' } },
      inputs: [[input]],
    }],
  };
}

async function saveHash(handle: WorkbookHandle): Promise<Uint8Array> {
  return new Uint8Array(await crypto.subtle.digest('SHA-256', new Uint8Array(handle.save()).buffer));
}

describe('workbook calculation context', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it.each([
    ['east of UTC', Date.UTC(2026, 0, 1, 23, 30), -120, Date.UTC(2026, 0, 2, 1, 30)],
    ['west of UTC', Date.UTC(2026, 0, 2, 2), 300, Date.UTC(2026, 0, 1, 21)],
    ['half-hour offset', Date.UTC(2026, 0, 1, 20), -330, Date.UTC(2026, 0, 2, 1, 30)],
  ] as const)('uses local NOW and TODAY for ordinary opens and cleared contexts (%s)', (_, ms, offset, localMs) => {
    const clock = spyOn(Date, 'now').mockReturnValue(ms);
    const timezone = spyOn(Date.prototype, 'getTimezoneOffset').mockReturnValue(offset);
    let handle: WorkbookHandle | undefined;
    const nowSerial = localMs / 86_400_000 + 25_569;
    try {
      handle = openWorkbook(volatileBytes());
      expect(values(handle).slice(0, 2)).toEqual([
        { kind: 'number', value: nowSerial },
        { kind: 'number', value: Math.floor(nowSerial) },
      ]);
      handle.setCalculationContext(CONTEXT);
      handle.editCell(0, 39, 25, '1');
      expect(values(handle).slice(0, 2)).toEqual([
        { kind: 'number', value: CONTEXT.nowSerial },
        { kind: 'number', value: Math.floor(CONTEXT.nowSerial) },
      ]);
      clock.mockReturnValue(ms + 86_400_000);
      handle.setCalculationContext(null);
      handle.editCell(0, 39, 25, '2');
      expect(values(handle).slice(0, 2)).toEqual([
        { kind: 'number', value: nowSerial + 1 },
        { kind: 'number', value: Math.floor(nowSerial) + 1 },
      ]);
    } finally {
      try { handle?.dispose(); }
      finally {
        timezone.mockRestore();
        clock.mockRestore();
      }
    }
  });

  it('opens and replays identical calls to identical values and SHA-256 save hashes', async () => {
    const source = volatileBytes();
    const first = openWorkbook(source, { calculation: CONTEXT });
    const second = openWorkbook(source, { calculation: CONTEXT });
    try {
      expect(values(first).slice(0, 2)).toEqual([
        { kind: 'number', value: CONTEXT.nowSerial },
        { kind: 'number', value: Math.floor(CONTEXT.nowSerial) },
      ]);
      const initial = values(first);
      const calls: Array<(handle: WorkbookHandle) => void> = [
        () => {},
        (handle) => { handle.editCell(0, 39, 0, '=NOW()+RANDBETWEEN(1,1000000)'); },
        (handle) => {
          handle.editCells(0, [
            { row: 39, col: 1, input: '=TODAY()' },
            { row: 39, col: 2, input: '=RANDBETWEEN(1,1000000)' },
          ]);
        },
        (handle) => { expect(handle.applyEdits(request(handle, '=RANDBETWEEN(1,1000000)+NOW()')).ok).toBe(true); },
        (handle) => {
          handle.applyOps([{ type: 'setRowHeight', sheet: 0, row: 39, height: 25 }]);
        },
        (handle) => { handle.editCellProfiled(0, 40, 0, '=RANDBETWEEN(1,1000000)'); },
        (handle) => {
          handle.applyOpsProfiled([{ type: 'setColWidth', sheet: 0, col: 8, width: 18 }]);
        },
        (handle) => { handle.patchRangeStyle(0, 'A40:C40', { bold: true }); },
        (handle) => { handle.setNumberFormat(0, 'A40:C40', 'number'); },
        (handle) => { handle.applyFormat(0, 'D40:F40', handle.captureFormat(0, 'A40:C40')); },
        (handle) => { handle.undo(); },
        (handle) => { handle.redo(); },
        (handle) => { handle.setActiveSheet(1); },
        (handle) => { handle.displayList(VIEWPORT); },
      ];
      for (const call of calls) {
        call(first);
        call(second);
        expect(values(first)).toEqual(initial);
        expect(values(second)).toEqual(initial);
        expect(await saveHash(first)).toEqual(await saveHash(second));
      }
      const reopened = openWorkbook(first.save(), { calculation: CONTEXT });
      try {
        expect(values(reopened)).toEqual(initial);
      } finally {
        reopened.dispose();
      }
    } finally {
      first.dispose();
      second.dispose();
    }
  });

  it('preserves shared draws when earlier independent random cells are inserted in different orders', () => {
    const source = volatileBytes();
    const first = openWorkbook(source, { calculation: CONTEXT });
    const second = openWorkbook(source, { calculation: CONTEXT });
    try {
      const initial = values(first);
      for (const value of initial.slice(2)) {
        expect(value).toMatchObject({ kind: 'number' });
      }
      expect(values(second)).toEqual(initial);
      for (const [handle, cols] of [[first, [0, 1]], [second, [1, 0]]] as const) {
        for (const col of cols) {
          expect(handle.editCell(0, 29, col, '=RANDBETWEEN(1,1000000)').applied).toBe(true);
          expect(values(handle)).toEqual(initial);
        }
      }
      expect(values(first)).toEqual(values(second));
      for (const handle of [first, second]) {
        const reopened = openWorkbook(handle.save(), { calculation: CONTEXT });
        try {
          expect(values(reopened)).toEqual(initial);
        } finally {
          reopened.dispose();
        }
      }
    } finally {
      first.dispose();
      second.dispose();
    }
  });

  it('keeps chart edits, undo and redo deterministic', async () => {
    const source = volatileBytes('charts');
    const first = openWorkbook(source, { calculation: CONTEXT });
    const second = openWorkbook(source, { calculation: CONTEXT });
    try {
      const chart = first.displayList(VIEWPORT).charts?.find((chart) => chart.movable);
      if (!chart) throw new Error('fixture must contain a movable chart');
      const initial = values(first);
      for (const call of [
        (handle: WorkbookHandle) => handle.moveChart(0, chart.id, 24, 12),
        (handle: WorkbookHandle) => handle.undo(),
        (handle: WorkbookHandle) => handle.redo(),
      ]) {
        expect(call(first).applied).toBe(true);
        expect(call(second).applied).toBe(true);
        expect(values(first)).toEqual(initial);
        expect(values(second)).toEqual(initial);
        expect(await saveHash(first)).toEqual(await saveHash(second));
      }
    } finally {
      first.dispose();
      second.dispose();
    }
  });

  it('uses a different seed for a different random value at the same clock', () => {
    const source = volatileBytes();
    const first = openWorkbook(source, { calculation: { ...CONTEXT, randSeed: 0 } });
    const second = openWorkbook(source, { calculation: { ...CONTEXT, randSeed: 0xffff_ffff } });
    try {
      expect(values(first).slice(0, 2)).toEqual(values(second).slice(0, 2));
      expect(values(first)[2]).not.toEqual(values(second)[2]);
      expect(values(first)[3]).not.toEqual(values(second)[3]);
    } finally {
      first.dispose();
      second.dispose();
    }
  });

  it('changes the handle context and restores the default clock and random path with null', () => {
    const handle = openWorkbook(volatileBytes());
    try {
      handle.setCalculationContext(CONTEXT);
      handle.editCell(0, 39, 25, '1');
      const pinned = values(handle);
      expect(pinned[0]).toEqual({ kind: 'number', value: CONTEXT.nowSerial });
      handle.editCell(0, 39, 25, '2');
      expect(values(handle)).toEqual(pinned);
      handle.setCalculationContext(null);
      const before = localSerial(Date.now());
      handle.editCell(0, 39, 25, '3');
      const after = localSerial(Date.now());
      const now = values(handle)[0] as { kind: string; value: number };
      expect(now.kind).toBe('number');
      expect(now.value).toBeGreaterThanOrEqual(before);
      expect(now.value).toBeLessThanOrEqual(after);
      const draws = [values(handle)[3]];
      for (const input of ['4', '5', '6']) {
        handle.editCell(0, 39, 25, input);
        draws.push(values(handle)[3]);
      }
      expect(draws.some((draw) => JSON.stringify(draw) !== JSON.stringify(pinned[3]))).toBe(true);
      expect(handle.applyEdits(request(handle, '7')).ok).toBe(true);
      expect(values(handle)[0]).toMatchObject({ kind: 'error' });
    } finally {
      handle.dispose();
    }
  });

  it('overrides only the clock for one request and resumes the handle context afterwards', () => {
    const source = volatileBytes();
    const handle = openWorkbook(source, { calculation: CONTEXT });
    const override = { nowSerial: 46_000.25 };
    const expected = openWorkbook(source, { calculation: { ...CONTEXT, ...override } });
    try {
      const initial = values(handle);
      expect(handle.applyEdits(request(handle, '1', override)).ok).toBe(true);
      expect(values(handle)).toEqual(values(expected));
      expect(values(handle).slice(2)).toEqual(initial.slice(2));
      handle.editCell(0, 39, 25, '2');
      expect(values(handle)).toEqual(initial);
      expect(handle.applyEdits(request(handle, '3', {})).ok).toBe(true);
      expect(values(handle)).toEqual(initial);
      handle.undo();
      expect(values(handle)).toEqual(initial);
    } finally {
      handle.dispose();
      expected.dispose();
    }
  });

  it('keeps values and save bytes identical for clock-only batch overrides with random formulas', () => {
    const source = volatileBytes();
    const first = openWorkbook(source, { calculation: CONTEXT });
    const second = openWorkbook(source, { calculation: CONTEXT });
    const override = { nowSerial: 46_000.25 };
    try {
      const initial = values(first);
      const draws: unknown[] = [];
      for (const handle of [first, second]) {
        expect(handle.applyEdits(request(handle, '=RANDBETWEEN(1,1000000)', override)))
          .toMatchObject({ ok: true, applied: true });
        expect(values(handle)[0]).toEqual({ kind: 'number', value: override.nowSerial });
        expect(values(handle).slice(2)).toEqual(initial.slice(2));
        const read = handle.readCells({
          ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'Z40' } }],
        });
        if (!read.ok) throw new Error(read.failure.message);
        const draw = read.ranges[0].cells[0][0].value;
        if (draw.kind !== 'number') throw new Error('Expected a random number');
        expect(draw.value).toBeGreaterThanOrEqual(1);
        expect(draw.value).toBeLessThanOrEqual(1_000_000);
        draws.push(draw);
      }
      expect(draws[0]).toEqual(draws[1]);
      expect(values(first)).toEqual(values(second));
      expect(first.save()).toEqual(second.save());
    } finally {
      first.dispose();
      second.dispose();
    }
  });

  it('leaves random streams unpinned when only the request clock is shared', () => {
    const source = volatileBytes();
    const first = openWorkbook(source);
    const second = openWorkbook(source);
    try {
      for (const handle of [first, second]) {
        expect(handle.applyEdits(request(handle, '=RANDBETWEEN(1,1000000)', {
          nowSerial: CONTEXT.nowSerial,
        }))).toMatchObject({ ok: true, applied: true });
        expect(values(handle).slice(0, 2)).toEqual([
          { kind: 'number', value: CONTEXT.nowSerial },
          { kind: 'number', value: Math.floor(CONTEXT.nowSerial) },
        ]);
      }
      expect(values(first).slice(2)).not.toEqual(values(second).slice(2));
      expect(first.save()).not.toEqual(second.save());
    } finally {
      first.dispose();
      second.dispose();
    }
  });

  it('rejects malformed contexts and edit overrides with TypeError', () => {
    const source = bytes();
    const handle = openWorkbook(source, { calculation: CONTEXT });
    try {
      for (const context of [
        null, undefined, 1, 'context', [], {},
        { nowSerial: 1 }, { randSeed: 1 },
        { nowSerial: NaN, randSeed: 1 }, { nowSerial: Infinity, randSeed: 1 },
        { nowSerial: -Infinity, randSeed: 1 }, { nowSerial: '1', randSeed: 1 },
        { nowSerial: 1, randSeed: NaN }, { nowSerial: 1, randSeed: Infinity },
        { nowSerial: 1, randSeed: -1 }, { nowSerial: 1, randSeed: 0.5 },
        { nowSerial: 1, randSeed: 0x1_0000_0000 }, { nowSerial: 1, randSeed: '1' },
        { ...CONTEXT, extra: true },
      ]) {
        if (context !== undefined) {
          expect(() => openWorkbook(source, { calculation: context as WorkbookCalculationContext })).toThrow(TypeError);
        }
        if (context !== null) {
          expect(() => handle.setCalculationContext(context as WorkbookCalculationContext)).toThrow(TypeError);
        }
      }
      for (const calculation of [
        null, [], 'context', { nowSerial: NaN }, { nowSerial: Infinity },
        { nowSerial: -Infinity }, { nowSerial: '1' }, { nowSerial: 1, extra: true },
        { randSeed: 7 }, { nowSerial: 1, randSeed: 7 },
        { randSeed: -1 }, { randSeed: 0.5 }, { randSeed: 0x1_0000_0000 }, { randSeed: '1' },
      ]) {
        const batch = { ...request(handle, '1'), calculation } as XlsxEditRequest;
        expect(() => handle.applyEdits(batch)).toThrow(TypeError);
        expect(() => handle.validateEdits(batch)).toThrow(TypeError);
      }
    } finally {
      handle.dispose();
    }
  });

  it('rejects contexts on collaborative handles at open and on the setter', () => {
    const source = bytes();
    expect(() => openWorkbook(source, { collaborative: true, calculation: CONTEXT })).toThrow(TypeError);
    const handle = openWorkbook(source, { collaborative: true, clientId: 731 });
    try {
      expect(() => handle.setCalculationContext(CONTEXT)).toThrow(TypeError);
      handle.setCalculationContext(null);
    } finally {
      handle.dispose();
    }
  });
});

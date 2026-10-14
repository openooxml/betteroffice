import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { WorkbookHandle, XlsxEditRequest, XlsxEditStep, XlsxRangeTarget } from '../index';
import { initWasm, openWorkbook } from './loader';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

function sampleBytes(): Uint8Array {
  return new Uint8Array(readFileSync(FIXTURE));
}

function target(a1: string, sheetId = 'sheet:0'): XlsxRangeTarget {
  return { sheetId, range: { kind: 'a1', a1 } };
}

function inputs(a1: string, values: string[][]): XlsxEditStep {
  return { op: 'setCellInputs', target: target(a1), inputs: values };
}

function batch(handle: WorkbookHandle, steps: XlsxEditStep[]): XlsxEditRequest {
  return { expectVersion: handle.version(), steps };
}

function cellValue(handle: WorkbookHandle, a1: string): unknown {
  const read = handle.readCells({ ranges: [target(a1)] });
  if (!read.ok) throw new Error(read.failure.message);
  return read.ranges[0].cells[0][0].value;
}

describe('xlsx edit batches', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('reads and finds with the version they were read at', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const version = handle.version();
      const read = handle.readCells({ ranges: [target('B3:D3')] });
      if (!read.ok) throw new Error(read.failure.message);
      expect(read.version).toBe(version);
      expect(read.sheets.map((sheet) => sheet.name)).toEqual(['Budget', 'Summary', 'Styled']);
      expect(read.ranges[0].cells[0][2]).toEqual({
        a1: 'D3',
        value: { kind: 'number', value: 157 },
        formula: 'B3+C3',
        displayText: '157',
      });
      const found = handle.findText({ text: 'Quarterly' });
      if (!found.ok) throw new Error(found.failure.message);
      expect(found.matches).toEqual([
        {
          cell: { sheetId: 'sheet:0', row: 0, col: 0, a1: 'A1' },
          text: 'Quarterly Budget Report',
        },
      ]);
      const limited = handle.findText({ text: 'Line item', limit: 2 });
      if (!limited.ok) throw new Error(limited.failure.message);
      expect(limited.matches.map((match) => match.cell.a1)).toEqual(['A3', 'A4']);
      expect(limited.truncated).toBe(true);
      const missing = handle.findText({ text: 'x', sheetIds: ['nope'] });
      expect(missing.ok).toBe(false);
      if (!missing.ok) expect(missing.failure.code).toBe('missing-target');
      expect(handle.version()).toBe(version);
    } finally {
      handle.dispose();
    }
  });

  it('applies a batch as one recalculated change that listeners see once', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const seen: Array<{ version: string; total: unknown }> = [];
      const unsubscribe = handle.onUpdate(() => {
        seen.push({ version: handle.version(), total: cellValue(handle, 'D3') });
      });
      const base = handle.version();
      const depth = handle.historyState().undoDepth;
      const result = handle.applyEdits(
        batch(handle, [
          inputs('B3', [['1000']]),
          { op: 'patchStyle', target: target('B3'), patch: { bold: true } },
          { op: 'setNumberFormat', target: target('C3'), format: 'percent' },
        ])
      );
      unsubscribe();
      if (!result.ok) throw new Error(result.failure.message);
      expect(result.applied).toBe(true);
      expect(result.baseVersion).toBe(base);
      expect(result.version).toBe(handle.version());
      expect(result.version).not.toBe(base);
      expect(result.changedSheets).toEqual(['sheet:0']);
      expect(result.calculation.truncated).toBe(false);
      expect(result.calculation.changed.map((cell) => cell.a1)).toContain('D3');
      expect(result.receipts.map((receipt) => receipt.changed)).toEqual([true, true, true]);
      expect(seen).toEqual([{ version: result.version, total: { kind: 'number', value: 1057 } }]);
      expect(handle.historyState().undoDepth).toBe(depth + 1);
      handle.undo();
      expect(handle.cell(0, 2, 1).input).toBe('100');
    } finally {
      handle.dispose();
    }
  });

  it('refuses as data with the workbook untouched', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      let updates = 0;
      handle.onUpdate(() => {
        updates += 1;
      });
      const version = handle.version();
      const guarded = handle.applyEdits(
        batch(handle, [
          inputs('B3', [['1']]),
          {
            ...inputs('B4', [['2']]),
            expect: { cells: [[{ value: { kind: 'number', value: 7 } }]] },
          },
        ])
      );
      expect(guarded.ok).toBe(false);
      if (!guarded.ok) {
        expect(guarded.version).toBe(version);
        expect(guarded.failure.code).toBe('content-mismatch');
        expect(guarded.failure.stepIndex).toBe(1);
        expect(guarded.failure.target).toEqual(target('B4'));
      }
      expect(handle.cell(0, 2, 1).input).toBe('100');
      expect(updates).toBe(0);

      handle.editCell(0, 9, 0, 'typed');
      const stale = handle.applyEdits({ expectVersion: version, steps: [inputs('B3', [['1']])] });
      expect(stale.ok).toBe(false);
      if (!stale.ok) expect(stale.failure.code).toBe('stale-version');

      const locked = handle.validateEdits(
        batch(handle, [inputs('B3:B4', [['1']])])
      );
      expect(locked.ok).toBe(false);
      if (!locked.ok) expect(locked.failure.code).toBe('invalid-step');

      expect(() =>
        handle.applyEdits({ steps: [] } as unknown as XlsxEditRequest)
      ).toThrow();
      expect(() =>
        handle.applyEdits(
          batch(handle, [{ op: 'insertRows' } as unknown as XlsxEditStep])
        )
      ).toThrow();
    } finally {
      handle.dispose();
    }
  });

  it('scopes versions to one workbook, whatever its client id', () => {
    const bytes = sampleBytes();
    const handles = [
      openWorkbook(bytes, { collaborative: true, clientId: 1 }),
      openWorkbook(bytes, { collaborative: true, clientId: 1 }),
      openWorkbook(bytes, { collaborative: true, clientId: 4_294_967_297 }),
    ];
    try {
      const versions = handles.map((handle) => handle.version());
      expect(new Set(versions).size).toBe(3);
      const [first, ...others] = handles;
      for (const other of others) {
        const foreign = first.applyEdits(batch(other, [inputs('B3', [['1']])]));
        expect(foreign.ok).toBe(false);
        if (!foreign.ok) expect(foreign.failure.code).toBe('stale-version');
      }
    } finally {
      for (const handle of handles) handle.dispose();
    }
  });

  it('validates without reserving anything and keeps history out on request', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const request = batch(handle, [inputs('B3', [['9']])]);
      const validated = handle.validateEdits(request);
      if (!validated.ok) throw new Error(validated.failure.message);
      expect(validated.wouldApply).toBe(true);
      expect(validated.previews).toEqual([
        { stepIndex: 0, target: target('B3'), wouldChange: true, changedCellCount: 1 },
      ]);
      expect(handle.version()).toBe(request.expectVersion);
      const history = handle.historyState();
      const quiet = handle.applyEdits({ ...request, history: 'none', source: 'agent' });
      if (!quiet.ok) throw new Error(quiet.failure.message);
      expect(quiet.source).toBe('agent');
      expect(handle.historyState()).toEqual(history);
      expect(handle.cell(0, 2, 1).input).toBe('9');
    } finally {
      handle.dispose();
    }
  });

  it('round-trips through save and reopen and converges on a peer', () => {
    const bytes = sampleBytes();
    const local = openWorkbook(bytes, { collaborative: true, clientId: 7101 });
    const peer = openWorkbook(bytes, { collaborative: true, clientId: 7102 });
    try {
      local.onUpdate((update) => peer.applyUpdate(update));
      const peerVersion = peer.version();
      const result = local.applyEdits({
        ...batch(local, [
          { op: 'setFormulas', target: target('E3'), formulas: [['D3*2']] },
          inputs('A10', [['added']]),
        ]),
        calculation: { nowSerial: 45000 },
      });
      if (!result.ok) throw new Error(result.failure.message);
      expect(peer.version()).not.toBe(peerVersion);
      expect(peer.cell(0, 2, 4).input).toBe('=D3*2');
      expect(cellValue(peer, 'E3')).toEqual({ kind: 'number', value: 314 });

      const reopened = openWorkbook(local.save());
      try {
        expect(reopened.cell(0, 2, 4).input).toBe('=D3*2');
        expect(reopened.cell(0, 9, 0).input).toBe('added');
        const stale = reopened.applyEdits({
          expectVersion: result.version,
          steps: [inputs('A10', [['again']])],
        });
        expect(stale.ok).toBe(false);
      } finally {
        reopened.dispose();
      }
    } finally {
      local.dispose();
      peer.dispose();
    }
  });
});

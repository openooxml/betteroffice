// Normalization: ZIP container headers, entry order, directory entries and compression only.
// These can differ between Rust targets; every part name and uncompressed byte must match.
// Proposal adapters rename fields and derive coordinates from A1; every wasm proposal field and native input is checked.
// Both opens recalculate with nowSerial=45000.75; the fixture and edits contain no volatile functions.
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, test } from 'bun:test';
import {
  initWasm,
  openWorkbook as openWasm,
  type NumberFormatMutation,
  type RangeStylePatch,
  type WorkbookHandle
} from '../../packages/xlsx/src/wasm/loader';
import { expectSavedPartsEqual } from '../native-parity';
import { openWorkbook, type Workbook } from './index.js';

const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url));
const nowSerial = 45000.75;

beforeAll(async () => {
  await initWasm(readFileSync(new URL('../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
});

function nativeEdits(workbook: Workbook) {
  return {
    set: (row: number, col: number, input: string) => workbook.set(0, `${String.fromCharCode(65 + col)}${row + 1}`, input, nowSerial),
    format: (range: string, format: NumberFormatMutation) => workbook.setNumberFormat(0, range, format, nowSerial),
    style: (range: string, patch: RangeStylePatch) => workbook.setStyle(0, range, patch, nowSerial),
    async proposal(numberFormat: NumberFormatMutation) {
      const proposal = await workbook.propose({
        agentId: 'parity-agent', note: 'formatted edit',
        edits: [{ sheet: 0, address: 'C4', input: '12.5', numberFormat }]
      }, nowSerial);
      expect(await workbook.proposals).toEqual([proposal]);
      expect(proposal.edits[0].input).toBe('12.5');
      const preview = { id: proposal.id, agentId: proposal.agentId, note: proposal.note, cells: proposal.edits.map((edit) => ({
        sheet: edit.sheet, row: Number(edit.address.slice(1)) - 1, col: edit.address.charCodeAt(0) - 65,
        a1: edit.address, oldText: edit.before, newText: edit.after
      })) };
      expect((await workbook.acceptProposal(proposal.id, false, nowSerial)).applied).toBe(true);
      expect(await workbook.proposals).toEqual([]);
      expect((await workbook.cell(0, 'C4')).input).toBe('12.5');
      return preview;
    }
  };
}

function wasmEdits(workbook: WorkbookHandle) {
  return {
    set: async (row: number, col: number, input: string) => workbook.editCell(0, row, col, input),
    format: async (range: string, format: NumberFormatMutation) => workbook.setNumberFormat(0, range, format),
    style: async (range: string, patch: RangeStylePatch) => workbook.patchRangeStyle(0, range, patch),
    async proposal(numberFormat: NumberFormatMutation) {
      const proposal = workbook.propose('parity-agent', 'formatted edit', [{ sheet: 0, row: 3, col: 2, input: '12.5', numberFormat }]);
      expect(workbook.listProposals()).toEqual([proposal]);
      expect(workbook.acceptProposal(proposal.id).applied).toBe(true);
      expect(workbook.listProposals()).toEqual([]);
      expect(workbook.cell(0, 3, 2).input).toBe('12.5');
      return proposal;
    }
  };
}

type Edits = {
  set(row: number, col: number, input: string): Promise<unknown>;
  format(range: string, format: NumberFormatMutation): Promise<unknown>;
  style(range: string, patch: RangeStylePatch): Promise<unknown>;
  proposal(numberFormat: NumberFormatMutation): Promise<unknown>;
};

async function compare(script: (edits: Edits) => Promise<unknown>) {
  const native = await openWorkbook(fixture, { recalculate: true, nowSerial });
  const wasm = openWasm(fixture, { calculation: { nowSerial, randSeed: 42 } });
  try {
    expect(await script(nativeEdits(native))).toEqual(await script(wasmEdits(wasm)));
    await expectSavedPartsEqual(await native.save(), wasm.save());
  } finally {
    wasm.dispose();
  }
}

describe('xlsx native/wasm saved parity', () => {
  test('plainText and plain cell edits', async () => {
    await compare(async (edits) => {
      await edits.set(0, 0, '42');
      await edits.set(0, 1, '=1+2');
      await edits.set(1, 0, 'Parity');
      await edits.format('A1', 'plainText');
    });
  });

  test('currency, date, percent, custom and decimal mutations', async () => {
    await compare(async (edits) => {
      for (const [col, format] of ['currency', 'date', 'percent'].entries()) {
        await edits.set(0, col, '123.456');
        await edits.format(`${String.fromCharCode(65 + col)}1`, format as NumberFormatMutation);
      }
      await edits.set(0, 3, '12.345');
      await edits.format('D1', { type: 'custom', pattern: '0.000" units"' });
      await edits.format('A1', 'increaseDecimal');
      await edits.format('C1', 'decreaseDecimal');
    });
  });

  test('borders, font, fill, alignment, wrapping and clear', async () => {
    await compare(async (edits) => {
      await edits.set(0, 0, 'styled');
      await edits.style('A1:B2', {
        bold: true, italic: true, strikethrough: true, fontFamily: 'Arial', fontSize: 16,
        textColor: '#123456', fillColor: '#abcdef',
        border: { preset: 'outer', style: 'dashed', color: '#654321' },
        horizontalAlignment: 'center', verticalAlignment: 'middle', textWrapping: 'wrap'
      });
      await edits.style('A1', { clear: ['bold', 'fillColor', 'borders', 'horizontalAlignment', 'textWrapping'] });
    });
  });

  test('proposal number formats and accepted edits', async () => {
    await compare(async (edits) => {
      await edits.set(3, 2, '7');
      return edits.proposal('currency');
    });
  });
});

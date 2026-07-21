/**
 * End-to-end proof of the wasm boundary without a browser.
 */

import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  StaleProposalError,
  initWasm,
  isPngExportAvailable,
  isProposalsAvailable,
  isWasmAvailable,
  openWorkbook,
  wasmVersion,
} from './loader';

// png files start with this 8-byte signature.
const PNG_MAGIC = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

// the committed, hand-built fixture (regenerate via e2e/fixtures/generate-sample.mjs).
const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

function sampleBytes(): Uint8Array {
  return new Uint8Array(readFileSync(FIXTURE));
}

describe('wasm loader', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('reports available and a version', () => {
    expect(isWasmAvailable()).toBe(true);
    expect(wasmVersion().length).toBeGreaterThan(0);
  });

  it('opens the hand-built fixture and reads sheet info', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const info = handle.sheetInfo();
      expect(info.sheetNames).toEqual(['Budget', 'Summary', 'Styled']);
      expect(info.activeSheet).toBe(0);
      expect(info.contentWidth).toBeGreaterThan(0);
      expect(info.contentHeight).toBeGreaterThan(0);
      expect(handle.calculationStatus()).toEqual({ limitedCells: [] });
    } finally {
      handle.dispose();
    }
  });

  it('renders a display list with real commands', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const dl = handle.displayList({ x: 0, y: 0, width: 400, height: 300 });
      expect(dl.width).toBe(400);
      expect(dl.commands.length).toBeGreaterThan(0);
      const texts = dl.commands.filter((c) => c.op === 'text');
      // the merged title is a shared string, so it must reach the display list.
      expect(texts.some((c) => c.op === 'text' && c.text.includes('Quarterly'))).toBe(true);
      // the formula cell's cached total value renders as a number command.
      expect(texts.some((c) => c.op === 'text' && c.text === '157')).toBe(true);
    } finally {
      handle.dispose();
    }
  });

  it('updates display-list font fields after style patch and format paint', () => {
    const handle = openWorkbook(sampleBytes());
    const viewport = { x: 0, y: 0, width: 500, height: 220 };
    try {
      const beforePatch = handle.displayList(viewport);
      const beforeSource = beforePatch.commands.find(
        (command) => command.op === 'text' && command.text === 'Line item 6'
      );
      expect(beforeSource).toMatchObject({ op: 'text', fontSize: 11 });

      handle.patchRangeStyle(0, 'A8', { bold: true, italic: true, fontFamily: 'Arial' });
      const afterPatch = handle.displayList(viewport);
      const afterSource = afterPatch.commands.find(
        (command) => command.op === 'text' && command.text === 'Line item 6'
      );
      expect(afterPatch).not.toEqual(beforePatch);
      expect(afterSource).toMatchObject({
        op: 'text',
        fontSize: 11,
        fontFamily: 'Arial',
        bold: true,
        italic: true,
      });

      const beforeApply = handle.displayList(viewport);
      handle.applyFormat(0, 'C8', handle.captureFormat(0, 'A8'));
      const afterApply = handle.displayList(viewport);
      const target = afterApply.commands.find(
        (command) => command.op === 'text' && command.text === '307' && command.clip?.y === 140
      );
      expect(afterApply).not.toEqual(beforeApply);
      expect(target).toMatchObject({
        op: 'text',
        fontSize: 11,
        fontFamily: 'Arial',
        bold: true,
        italic: true,
      });
    } finally {
      handle.dispose();
    }
  });

  it('switches the active sheet', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      handle.setActiveSheet(1);
      expect(handle.sheetInfo().activeSheet).toBe(1);
      expect(() => handle.setActiveSheet(9)).toThrow();
    } finally {
      handle.dispose();
    }
  });

  it('rejects non-workbook bytes with an Error', () => {
    expect(() => openWorkbook(new Uint8Array([1, 2, 3]))).toThrow();
  });

  it('renders the fixture viewport to png bytes', () => {
    expect(isPngExportAvailable()).toBe(true);
    const handle = openWorkbook(sampleBytes());
    try {
      const png = handle.renderPng({ x: 0, y: 0, width: 400, height: 300 });
      expect(png.length).toBeGreaterThan(100);
      expect(Array.from(png.subarray(0, 8))).toEqual(PNG_MAGIC);
    } finally {
      handle.dispose();
    }
  });

  // the wasm is inconsistent: editCell/editCells return a bare SheetInfo while
  // undo/redo return the {applied, sheetInfo} envelope. the loader normalizes
  // both, so callers always get an EditResult with a populated sheetInfo.
  it('normalizes edit/undo results to a populated EditResult', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const edit = handle.editCell(0, 7, 0, '123');
      expect(edit.applied).toBe(true);
      expect(edit.sheetInfo.sheetNames).toEqual(['Budget', 'Summary', 'Styled']);
      expect(edit.sheetInfo.contentHeight).toBeGreaterThan(0);
      expect(handle.cell(0, 7, 0).input).toBe('123');

      const undone = handle.undo();
      expect(undone.applied).toBe(true);
      expect(undone.sheetInfo.sheetNames).toEqual(['Budget', 'Summary', 'Styled']);
    } finally {
      handle.dispose();
    }
  });

  it('round-trips formatting, format capture, merge metadata, and history state', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      expect(handle.historyState()).toEqual({
        canUndo: false,
        canRedo: false,
        undoDepth: 0,
        redoDepth: 0,
      });
      expect(handle.selectionFormatting(0, 'A8:B8').bold).toBe(false);
      expect(
        handle.patchRangeStyle(0, 'A8:B8', {
          bold: true,
          fontFamily: 'Arial',
          textColor: '#123456',
        }).applied
      ).toBe(true);
      expect(handle.selectionFormatting(0, 'A8:B8')).toMatchObject({
        bold: true,
        fontFamily: 'Arial',
        textColor: '#123456',
      });
      handle.setNumberFormat(0, 'A8:B8', 'percent');
      expect(handle.selectionFormatting(0, 'A8:B8').numberFormat).toBe('percent');
      handle.setNumberFormat(0, 'A8:B8', 'increaseDecimal');
      expect(handle.selectionFormatting(0, 'A8:B8').numberFormatPattern).toBe('0.000%');
      const captured = handle.captureFormat(0, 'A8');
      expect(captured).toMatchObject({ rows: 1, columns: 1 });
      handle.applyFormat(0, 'C8', captured);
      expect(handle.selectionFormatting(0, 'C8')).toMatchObject({
        bold: true,
        numberFormat: 'percent',
      });
      expect(handle.historyState()).toMatchObject({
        canUndo: true,
        canRedo: false,
        undoDepth: 4,
        redoDepth: 0,
      });
      handle.undo();
      expect(handle.historyState()).toMatchObject({ undoDepth: 3, redoDepth: 1 });
      expect(handle.mergedRanges(0, 'A1:Z20').length).toBeGreaterThan(0);
    } finally {
      handle.dispose();
    }
  });
});

// the two paths are gated on `isProposalsAvailable()` so this file passes
// whether or not the embedded module carries the proposals api: the full loop
// runs when it's present, the graceful-degrade path when it isn't. both are
// kept so a future core rebuild (either direction) needs no test rewrite.
describe('wasm loader — proposals', () => {
  const available = isProposalsAvailable();

  // runs against a core WITHOUT the api: it must be detectable-absent and the
  // handle must degrade without throwing on the read paths.
  it.skipIf(available)('reports unavailable and degrades gracefully on an old core', () => {
    expect(isProposalsAvailable()).toBe(false);
    const handle = openWorkbook(sampleBytes());
    try {
      expect(handle.isProposalsAvailable()).toBe(false);
      expect(handle.listProposals()).toEqual([]);
      expect(handle.rejectProposal('p1')).toBe(false);
      expect(() => handle.propose('demo-agent', 'note', [])).toThrow();
      expect(() => handle.acceptProposal('p1')).toThrow();
    } finally {
      handle.dispose();
    }
  });

  it.skipIf(!available)('stages, lists, and applies a proposal as one undo step', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      expect(handle.isProposalsAvailable()).toBe(true);

      // E7 is empty in the fixture body; propose a value there.
      const proposal = handle.propose('demo-agent', 'column totals', [
        { sheet: 0, row: 6, col: 4, input: '=SUM(B3:B4)' },
      ]);
      expect(proposal.agentId).toBe('demo-agent');
      expect(proposal.note).toBe('column totals');
      expect(proposal.cells).toHaveLength(1);
      expect(proposal.cells[0].a1).toBe('E7');
      // old/new are formatted display texts, not the raw input.
      expect(proposal.cells[0].newText).not.toContain('=SUM');

      // staging does not mutate the grid.
      expect(handle.cell(0, 6, 4).input).toBe('');
      expect(handle.listProposals().map((p) => p.id)).toContain(proposal.id);

      const accepted = handle.acceptProposal(proposal.id);
      expect(accepted.applied).toBe(true);
      expect(accepted.proposalId).toBe(proposal.id);
      expect(handle.cell(0, 6, 4).input).toBe('=SUM(B3:B4)');
      // accepting clears it from the pending list.
      expect(handle.listProposals().map((p) => p.id)).not.toContain(proposal.id);

      // one undo reverts the whole proposal.
      handle.undo();
      expect(handle.cell(0, 6, 4).input).toBe('');
    } finally {
      handle.dispose();
    }
  });

  it.skipIf(!available)('throws StaleProposalError when the base changed, force applies', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const proposal = handle.propose('demo-agent', null, [
        { sheet: 0, row: 6, col: 4, input: '10' },
      ]);
      // mutate the proposal's target cell out from under it.
      handle.editCell(0, 6, 4, '999');

      let stale: StaleProposalError | null = null;
      try {
        handle.acceptProposal(proposal.id);
      } catch (e) {
        stale = e instanceof StaleProposalError ? e : null;
      }
      expect(stale).toBeInstanceOf(StaleProposalError);
      expect(stale?.cells).toContain('E7');

      // force applies despite the drift.
      const forced = handle.acceptProposal(proposal.id, { force: true });
      expect(forced.applied).toBe(true);
      expect(handle.cell(0, 6, 4).input).toBe('10');
    } finally {
      handle.dispose();
    }
  });

  it.skipIf(!available)('rejects a proposal without applying it', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const proposal = handle.propose('demo-agent', null, [
        { sheet: 0, row: 6, col: 4, input: '42' },
      ]);
      expect(handle.rejectProposal(proposal.id)).toBe(true);
      expect(handle.listProposals().map((p) => p.id)).not.toContain(proposal.id);
      // the cell is untouched and a second reject is a no-op.
      expect(handle.cell(0, 6, 4).input).toBe('');
      expect(handle.rejectProposal(proposal.id)).toBe(false);
    } finally {
      handle.dispose();
    }
  });
});

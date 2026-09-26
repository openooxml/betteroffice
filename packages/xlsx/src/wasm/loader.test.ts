/**
 * End-to-end proof of the wasm boundary without a browser.
 */

import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { XlsxTextMatch, XlsxTextSearchOptions } from '../index';

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
const CHART_FIXTURE = resolve(import.meta.dir, '../../test-fixtures/charts.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

function sampleBytes(): Uint8Array {
  return new Uint8Array(readFileSync(FIXTURE));
}

function chartBytes(): Uint8Array {
  return new Uint8Array(readFileSync(CHART_FIXTURE));
}

describe('wasm loader', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('reports available and a version', () => {
    expect(isWasmAvailable()).toBe(true);
    expect(wasmVersion().length).toBeGreaterThan(0);
  });

  it('restores persisted state from the published 0.2.1 package', () => {
    const update = new Uint8Array(readFileSync(resolve(
      import.meta.dir,
      '../../../../crates/betteroffice-xlsx/tests/fixtures/workbook-npm-0.2.1.update.bin'
    )));
    const handle = openWorkbook(sampleBytes(), { collaborative: true, clientId: 5020 });
    const peer = openWorkbook(sampleBytes(), { collaborative: true, clientId: 5021 });
    try {
      handle.applyUpdate(update);
      expect(handle.cell(0, 42, 0).input).toBe('PublishedReleaseState');
      handle.editCell(0, 42, 1, 'after restore');
      peer.applyUpdate(handle.encodeStateAsUpdate());
      expect(peer.cell(0, 42, 0).input).toBe('PublishedReleaseState');
      expect(peer.cell(0, 42, 1).input).toBe('after restore');
      const reopened = openWorkbook(handle.save());
      try {
        expect(reopened.cell(0, 42, 0).input).toBe('PublishedReleaseState');
        expect(reopened.cell(0, 42, 1).input).toBe('after restore');
      } finally {
        reopened.dispose();
      }
    } finally {
      handle.dispose();
      peer.dispose();
    }
  });

  it('refuses an old snapshot after a local edit', () => {
    const update = new Uint8Array(readFileSync(resolve(
      import.meta.dir,
      '../../../../crates/betteroffice-xlsx/tests/fixtures/workbook-npm-0.2.1.update.bin'
    )));
    const handle = openWorkbook(sampleBytes(), { collaborative: true, clientId: 5022 });
    try {
      handle.editCell(0, 42, 0, 'local edit');
      expect(() => handle.applyUpdate(update)).toThrow();
      expect(handle.cell(0, 42, 0).input).toBe('local edit');
    } finally {
      handle.dispose();
    }
  });

  it('preserves a cleared cell and its undo history when applying a snapshot', () => {
    for (const legacy of [false, true]) {
      const handle = openWorkbook(sampleBytes(), { collaborative: true, clientId: 5023 });
      try {
        const original = handle.cell(0, 0, 0).input;
        expect(original).not.toBe('');
        const update = legacy
          ? new Uint8Array(readFileSync(resolve(
            import.meta.dir,
            '../../../../crates/betteroffice-xlsx/tests/fixtures/workbook-npm-0.2.1.update.bin'
          )))
          : handle.encodeStateAsUpdate();
        handle.editCell(0, 0, 0, '');
        expect(handle.cell(0, 0, 0).input).toBe('');
        if (legacy) {
          expect(() => handle.applyUpdate(update)).toThrow();
        } else {
          handle.applyUpdate(update);
        }
        expect(handle.cell(0, 0, 0).input).toBe('');
        handle.undo();
        expect(handle.cell(0, 0, 0).input).toBe(original);
      } finally {
        handle.dispose();
      }
    }
  });

  it('opens the hand-built fixture and reads sheet info', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      const info = handle.sheetInfo();
      expect(info.sheetIds).toEqual(['sheet:0', 'sheet:1', 'sheet:2']);
      expect(info.sheetNames).toEqual(['Budget', 'Summary', 'Styled']);
      expect(info.activeSheet).toBe(0);
      expect(info.contentWidth).toBeGreaterThan(0);
      expect(info.contentHeight).toBeGreaterThan(0);
      expect(info.frozenRows).toBe(0);
      expect(info.frozenCols).toBe(0);
      expect(info.initialScrollX).toBe(0);
      expect(info.initialScrollY).toBe(0);
      expect(handle.calculationStatus()).toEqual({ limitedCells: [] });
      const position = handle.cellPosition(0, 7, 2);
      expect(position.x).toBeGreaterThan(0);
      expect(position.y).toBeGreaterThan(0);
    } finally {
      handle.dispose();
    }
  });

  it('pages stored addresses through the public handle after peer edits and save', () => {
    const left = openWorkbook(sampleBytes(), { collaborative: true, clientId: 5030 });
    const right = openWorkbook(sampleBytes(), { collaborative: true, clientId: 5031 });
    try {
      left.editCell(0, 40, 10, 'sparse');
      left.patchRangeStyle(0, 'L42', { bold: true });
      right.applyUpdate(left.encodeStateAsUpdate());
      const stateBeforeRead = right.encodeStateVector();
      const addresses: string[] = [];
      let after: { row: number; col: number } | undefined;
      do {
        const page = right.storedCellAddresses(0, { after, limit: 3 });
        expect(page.cells.length).toBeGreaterThan(0);
        expect(page.cells.length).toBeLessThanOrEqual(3);
        addresses.push(...page.cells.map((cell) => cell.a1));
        after = page.next ?? undefined;
      } while (after);
      expect(addresses).toContain('K41');
      expect(addresses).toContain('L42');
      expect(new Set(addresses).size).toBe(addresses.length);
      expect(right.encodeStateVector()).toEqual(stateBeforeRead);
      expect(() => right.storedCellAddresses(0, { limit: 0 })).toThrow(RangeError);
      expect(() => right.storedCellAddresses(0, { limit: 10_001 })).toThrow(RangeError);
      expect(() => right.storedCellAddresses(0, { after: { row: 1_048_576, col: 0 } })).toThrow();

      const reopened = openWorkbook(right.save());
      try {
        const saved = reopened.storedCellAddresses(0, { after: { row: 40, col: 9 }, limit: 100 });
        expect(saved.cells[0]?.a1).toBe('K41');
        expect(saved.cells.map((cell) => cell.a1)).toContain('L42');
      } finally {
        reopened.dispose();
      }
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  it('restores inherited column formatting when undoing a deletion', () => {
    const bytes = new Uint8Array(readFileSync(resolve(
      import.meta.dir,
      '../../../../crates/betteroffice-xlsx/tests/fixtures/column-style-undo.xlsx'
    )));
    const handle = openWorkbook(bytes);
    const viewport = { x: 0, y: 0, width: 500, height: 250 };
    const redFills = () => JSON.stringify(handle.displayList(viewport)).match(/#ff0000/gi)?.length ?? 0;
    try {
      expect(redFills()).toBeGreaterThan(0);
      const before = handle.displayList(viewport);
      handle.applyOps([{ type: 'deleteCols', sheet: 0, at: 1, count: 1 }]);
      expect(redFills()).toBe(0);
      for (let i = 0; i < 2; i += 1) {
        handle.undo();
        expect(handle.displayList(viewport)).toEqual(before);
        const reopened = openWorkbook(handle.save());
        try {
          expect(reopened.displayList(viewport)).toEqual(before);
        } finally {
          reopened.dispose();
        }
        handle.redo();
        expect(redFills()).toBe(0);
      }
    } finally {
      handle.dispose();
    }
  });

  it('searches formatted cell text without changing workbook state', () => {
    const handle = openWorkbook(sampleBytes());
    try {
      handle.editCell(0, 20, 1, 'QueryNeedle');
      handle.editCell(1, 0, 0, 'queryneedle');
      handle.editCell(2, 4, 2, '0.25');
      handle.setNumberFormat(2, 'C5', 'percent');
      const before = handle.encodeStateVector();

      const options: XlsxTextSearchOptions = { caseSensitive: false };
      const matches: XlsxTextMatch[] = handle.searchText('queryneedle', options);
      expect(matches).toEqual([
        {
          sheet: 0,
          sheetId: 'sheet:0',
          sheetName: 'Budget',
          row: 20,
          col: 1,
          a1: 'B21',
          text: 'QueryNeedle',
        },
        {
          sheet: 1,
          sheetId: 'sheet:1',
          sheetName: 'Summary',
          row: 0,
          col: 0,
          a1: 'A1',
          text: 'queryneedle',
        },
      ]);
      expect(handle.searchText('QueryNeedle', { caseSensitive: true })).toHaveLength(1);
      expect(handle.searchText('queryneedle', { limit: 1 })).toHaveLength(1);
      expect(handle.searchText('25.00%')[0]).toMatchObject({
        sheet: 2,
        a1: 'C5',
        text: '25.00%',
      });
      expect(handle.searchText('')).toEqual([]);
      expect(() => handle.searchText('queryneedle', { limit: -1 })).toThrow(RangeError);
      expect(handle.encodeStateVector()).toEqual(before);
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

  it('renders every chart anchor variant through the wasm display list', () => {
    const handle = openWorkbook(chartBytes());
    try {
      const dl = handle.displayList({ x: 0, y: 0, width: 800, height: 800 });
      expect(dl.charts?.length).toBe(4);
      expect(dl.commands.some((command) => command.op === 'path')).toBe(true);
      expect(
        dl.commands.some(
          (command) => command.op === 'text' && command.chart && command.text === 'Revenue trend'
        )
      ).toBe(true);
    } finally {
      handle.dispose();
    }
  });

  it('resolves a point to the chart under it and nothing outside one', () => {
    const handle = openWorkbook(chartBytes());
    const viewport = { x: 0, y: 0, width: 800, height: 800 };
    try {
      const chart = (handle.displayList(viewport).charts ?? [])[0];
      expect(chart.id.length).toBeGreaterThan(0);
      const hit = handle.chartAtPoint(
        viewport,
        chart.clip.x + chart.clip.w / 2,
        chart.clip.y + chart.clip.h / 2
      );
      expect(hit?.id).toBe(chart.id);
      expect(hit?.rect).toEqual(chart.rect);
      expect(handle.chartAtPoint(viewport, chart.rect.x - 8, chart.rect.y - 8)).toBeNull();
    } finally {
      handle.dispose();
    }
  });

  // a chart part backs a frame, not the other way round: two anchors may name
  // one part, so only the anchor addresses a frame across the boundary.
  it('addresses a chart by its drawing anchor and moves that one alone', () => {
    const handle = openWorkbook(chartBytes());
    const viewport = { x: 0, y: 0, width: 800, height: 800 };
    try {
      const before = handle.displayList(viewport).charts ?? [];
      expect(before.map((chart) => chart.id)).toEqual([
        'xl/drawings/drawing1.xml#0',
        'xl/drawings/drawing1.xml#1',
        'xl/drawings/drawing1.xml#2',
        'xl/drawings/drawing1.xml#3',
      ]);

      const target = before.findIndex((chart) => chart.movable);
      expect(handle.moveChart(0, before[target].id, 24, 12).applied).toBe(true);
      const after = handle.displayList(viewport).charts ?? [];
      expect(after.map((chart) => chart.id)).toEqual(before.map((chart) => chart.id));
      for (const [index, chart] of after.entries()) {
        if (index === target) expect(chart.rect).not.toEqual(before[index].rect);
        else expect(chart.rect).toEqual(before[index].rect);
      }
    } finally {
      handle.dispose();
    }
  });

  it('moves a chart by a pixel delta that survives save and undoes in one step', () => {
    const handle = openWorkbook(chartBytes());
    const viewport = { x: 0, y: 0, width: 800, height: 800 };
    try {
      const before = (handle.displayList(viewport).charts ?? [])[0];
      expect(handle.moveChart(0, before.id, 24, 12).applied).toBe(true);

      const moved = (handle.displayList(viewport).charts ?? [])[0];
      expect(moved.rect.x).toBeCloseTo(before.rect.x + 24, 1);
      expect(moved.rect.y).toBeCloseTo(before.rect.y + 12, 1);
      expect(moved.rect.w).toBeCloseTo(before.rect.w, 1);
      expect(moved.rect.h).toBeCloseTo(before.rect.h, 1);

      const reopened = openWorkbook(handle.save());
      try {
        expect((reopened.displayList(viewport).charts ?? [])[0].rect).toEqual(moved.rect);
      } finally {
        reopened.dispose();
      }

      handle.undo();
      expect((handle.displayList(viewport).charts ?? [])[0].rect).toEqual(before.rect);
    } finally {
      handle.dispose();
    }
  });

  it('refuses to move a chart pinned to the sheet by an absolute anchor', () => {
    const handle = openWorkbook(chartBytes());
    const viewport = { x: 0, y: 0, width: 800, height: 800 };
    try {
      const charts = handle.displayList(viewport).charts ?? [];
      const pinned = charts.filter((chart) => !chart.movable);
      expect(pinned.length).toBe(1);
      expect(() => handle.moveChart(0, pinned[0].id, 1, 0)).toThrow(/pinned/);
      expect(() => handle.moveChart(0, 'xl/drawings/drawing1.xml#9', 1, 0)).toThrow(
        /drawing1\.xml#9/
      );
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
        { sheet: 0, row: 6, col: 4, input: '0.5', numberFormat: 'percent' },
      ]);
      expect(proposal.agentId).toBe('demo-agent');
      expect(proposal.note).toBe('column totals');
      expect(proposal.cells).toHaveLength(1);
      expect(proposal.cells[0].a1).toBe('E7');
      // old/new are formatted display texts, not the raw input.
      expect(proposal.cells[0].newText).toBe('50.00%');

      // staging does not mutate the grid.
      expect(handle.cell(0, 6, 4).input).toBe('');
      expect(handle.listProposals().map((p) => p.id)).toContain(proposal.id);

      const accepted = handle.acceptProposal(proposal.id);
      expect(accepted.applied).toBe(true);
      expect(accepted.proposalId).toBe(proposal.id);
      expect(handle.cell(0, 6, 4).input).toBe('0.5');
      expect(handle.selectionFormatting(0, 'E7').numberFormat).toBe('percent');
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

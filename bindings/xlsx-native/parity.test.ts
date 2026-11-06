// Normalization: ZIP container headers, entry order, directory entries and compression only.
// These can differ between Rust targets; every part name and uncompressed byte must match.
// Both opens recalculate with nowSerial=45000.75 and randSeed=42; the fixture and edits contain no volatile functions.
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, test } from 'bun:test';
import * as wasm from '../../packages/xlsx/src/wasm/loader';
import { FORMATS } from '../../scripts/api-surface';
import { expectSavedPartsEqual } from '../native-parity';

type Xlsx = typeof wasm;
type WorkbookHandle = wasm.WorkbookHandle;

const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url));
const calculation = { nowSerial: 45000.75, randSeed: 42 };
let native: Xlsx;

beforeAll(async () => {
  native = await import('./dist/index.js') as unknown as Xlsx;
  await native.initWasm();
  await wasm.initWasm(readFileSync(new URL('../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
});

async function compare(script: (workbook: WorkbookHandle) => unknown) {
  const nativeWorkbook = native.openWorkbook(fixture, { calculation });
  const wasmWorkbook = wasm.openWorkbook(fixture, { calculation });
  try {
    expect(script(nativeWorkbook)).toEqual(script(wasmWorkbook));
    await expectSavedPartsEqual(nativeWorkbook.save(), wasmWorkbook.save());
  } finally {
    nativeWorkbook.dispose();
    wasmWorkbook.dispose();
  }
}

describe('xlsx native/wasm package parity', () => {
  test('each entry exports the names of the matching @betteroffice/xlsx entry', async () => {
    const manifest = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));
    expect(Object.keys(FORMATS.xlsx.entries)).toEqual(Object.keys(manifest.exports));
    for (const [entry, source] of Object.entries(FORMATS.xlsx.entries)) {
      const nativeKeys = Object.keys(await import(new URL(manifest.exports[entry].import, import.meta.url).href));
      const wasmKeys = Object.keys(await import(`../../packages/xlsx/${source}`));
      expect(nativeKeys.sort()).toEqual(wasmKeys.sort());
    }
  });

  test('plainText and plain cell edits', async () => {
    await compare((workbook) => [
      workbook.editCell(0, 0, 0, '42'),
      workbook.editCell(0, 0, 1, '=1+2'),
      workbook.editCell(0, 1, 0, 'Parity'),
      workbook.setNumberFormat(0, 'A1', 'plainText'),
    ]);
  });

  test('currency, date, percent, custom and decimal mutations', async () => {
    await compare((workbook) => {
      const results: unknown[] = [];
      for (const [col, format] of (['currency', 'date', 'percent'] as const).entries()) {
        results.push(workbook.editCell(0, 0, col, '123.456'));
        results.push(workbook.setNumberFormat(0, `${String.fromCharCode(65 + col)}1`, format));
      }
      results.push(workbook.editCell(0, 0, 3, '12.345'));
      results.push(workbook.setNumberFormat(0, 'D1', { type: 'custom', pattern: '0.000" units"' }));
      results.push(workbook.setNumberFormat(0, 'A1', 'increaseDecimal'));
      results.push(workbook.setNumberFormat(0, 'C1', 'decreaseDecimal'));
      return results;
    });
  });

  test('borders, font, fill, alignment, wrapping and clear', async () => {
    await compare((workbook) => [
      workbook.editCell(0, 0, 0, 'styled'),
      workbook.patchRangeStyle(0, 'A1:B2', {
        bold: true, italic: true, strikethrough: true, fontFamily: 'Arial', fontSize: 16,
        textColor: '#123456', fillColor: '#abcdef',
        border: { preset: 'outer', style: 'dashed', color: '#654321' },
        horizontalAlignment: 'center', verticalAlignment: 'middle', textWrapping: 'wrap'
      }),
      workbook.patchRangeStyle(0, 'A1', { clear: ['bold', 'fillColor', 'borders', 'horizontalAlignment', 'textWrapping'] }),
    ]);
  });

  test('proposal number formats and accepted edits', async () => {
    await compare((workbook) => {
      workbook.editCell(0, 3, 2, '7');
      const proposal = workbook.propose('parity-agent', 'formatted edit', [{ sheet: 0, row: 3, col: 2, input: '12.5', numberFormat: 'currency' }]);
      const listed = workbook.listProposals();
      const accepted = workbook.acceptProposal(proposal.id);
      return [proposal, listed, accepted, workbook.listProposals(), workbook.cell(0, 3, 2)];
    });
  });

  test('structural edits retain rebased and stale proposals', async () => {
    await compare((workbook) => {
      const proposal = workbook.propose('agent', null, [{ sheet: 0, row: 6, col: 4, input: '42' }]);
      workbook.applyOps([{ type: 'insertRows', sheet: 0, at: 5, count: 2 }]);
      expect(workbook.listProposals()[0]?.cells[0]?.a1).toBe('E9');
      workbook.applyOps([{ type: 'deleteRows', sheet: 0, at: 8, count: 1 }]);
      expect(workbook.listProposals()[0]?.cells[0]).toMatchObject({ a1: 'E9', stale: true });
      expect(workbook.listProposals()[0]?.staleTargets).toEqual(['E9']);
      expect(() => workbook.acceptProposal(proposal.id)).toThrow();
      workbook.undo();
      const restored = workbook.listProposals();
      workbook.acceptProposal(proposal.id);
      return [restored, workbook.cell(0, 8, 4), workbook.listProposals()];
    });
  });

  test('cell reads, markdown and structured exports', async () => {
    await compare((workbook) => {
      const read = workbook.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:D4' } }] });
      if (!read.ok) throw new Error('Expected cell read');
      const { version, ...cells } = read;
      expect(version).toMatch(/^[0-9a-f]{32}-0$/);
      return [cells, workbook.sheetInfo()];
    });
    expect(await native.exportXlsxMarkdown(fixture)).toEqual(await wasm.exportXlsxMarkdown(fixture));
    expect(await native.exportXlsxStructured(fixture)).toEqual(await wasm.exportXlsxStructured(fixture));
  });

  test('PNG range targets refuse bad ranges with the same error as the wasm range path', async () => {
    const nativeWorkbook = native.openWorkbook(fixture, { calculation });
    const wasmWorkbook = wasm.openWorkbook(fixture, { calculation });
    const failure = (error: unknown) => ({ isError: error instanceof Error, name: (error as Error)?.name, message: String((error as Error)?.message ?? error) });
    try {
      for (const range of ['not a range', 'A0:B2', 'A1:XFE1', 'b2:', { kind: 'a1', a1: 'A1:B' }] as const) {
        const target = { sheet: 0, range } as Parameters<WorkbookHandle['renderPng']>[0];
        let wasmFailure: unknown;
        try { wasmWorkbook.renderPngFonts(target); } catch (error) { wasmFailure = failure(error); }
        let nativeFailure: unknown;
        try { await nativeWorkbook.renderPng(target); } catch (error) { nativeFailure = failure(error); }
        expect(wasmFailure).toMatchObject({ isError: true, message: expect.stringMatching(/^bad range: /) });
        expect(nativeFailure).toEqual(wasmFailure);
      }
    } finally {
      nativeWorkbook.dispose();
      wasmWorkbook.dispose();
    }
  });

  test('PNG range targets read lowercase A1 as Excel does, on both engines', async () => {
    const nativeWorkbook = native.openWorkbook(fixture, { calculation });
    const wasmWorkbook = wasm.openWorkbook(fixture, { calculation });
    try {
      for (const [lower, upper] of [['a1:b2', 'A1:B2'], ['c3', 'C3'], ['a1:D4', 'A1:D4']]) {
        const fonts = (workbook: WorkbookHandle, range: string) => workbook.renderPngFonts({ sheet: 0, range });
        expect(fonts(wasmWorkbook, lower)).toEqual(fonts(wasmWorkbook, upper));
        expect(fonts(nativeWorkbook, lower)).toEqual(fonts(wasmWorkbook, upper));
        expect(fonts(nativeWorkbook, upper)).toEqual(fonts(wasmWorkbook, upper));
        const png = async (range: string) => Buffer.from(await nativeWorkbook.renderPng({ sheet: 0, range }));
        expect((await png(lower)).equals(await png(upper))).toBe(true);
      }
    } finally {
      nativeWorkbook.dispose();
      wasmWorkbook.dispose();
    }
  });
});

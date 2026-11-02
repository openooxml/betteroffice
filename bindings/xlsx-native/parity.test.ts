// Normalization: ZIP container headers, entry order, directory entries and compression only.
// These can differ between Rust targets; every part name and uncompressed byte must match.
// Both opens recalculate with nowSerial=45000.75 and randSeed=42; the fixture and edits contain no volatile functions.
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
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

async function compare(script: (workbook: WorkbookHandle) => unknown, bytes: Uint8Array = fixture) {
  const nativeWorkbook = native.openWorkbook(bytes, { calculation });
  const wasmWorkbook = wasm.openWorkbook(bytes, { calculation });
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

  test('locked cells on a protected sheet refuse cell edits', async () => {
    const zip = await JSZip.loadAsync(fixture);
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    zip.file('xl/worksheets/sheet1.xml', sheet.replace('</sheetData>', '</sheetData><sheetProtection sheet="1"/>'));
    const locked = await zip.generateAsync({ type: 'uint8array' });
    const refusal = (edit: () => unknown) => {
      try {
        edit();
        return null;
      } catch (error) {
        return (error as Error).message;
      }
    };
    await compare((workbook) => {
      const version = workbook.version();
      const refusals = [
        refusal(() => workbook.editCell(0, 2, 1, '1')),
        refusal(() => workbook.editCells(0, [{ row: 2, col: 1, input: '1' }])),
        refusal(() => workbook.patchRangeStyle(0, 'B3', { bold: true })),
      ];
      expect(refusals).toEqual([
        'B3 is locked on protected sheet "Budget"',
        'B3 is locked on protected sheet "Budget"',
        'sheet "Budget" is protected',
      ]);
      expect(workbook.version()).toBe(version);
      return [refusals, workbook.cell(0, 2, 1), workbook.editCell(1, 0, 0, 'open')];
    }, locked);
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
});

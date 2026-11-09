// Normalization: ZIP container headers, entry order, directory entries and compression only.
// These can differ between Rust targets; every part name and uncompressed byte must match.
// Both opens recalculate with nowSerial=45000.75 and randSeed=42; the fixture and edits contain no volatile functions.
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
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

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DOC_RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** `First`, `Data` holding table `Amounts` (1, 2), `Later` holding table `Prices` (10, 20), and `Other`
 * summing both through structured references. */
async function tablesFixture() {
  const rels = (entries: [string, string, string][]) =>
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries
      .map(([id, kind, target]) => `<Relationship Id="${id}" Type="${DOC_RELS}/${kind}" Target="${target}"/>`)
      .join('')}</Relationships>`;
  const sheet = (rows: string, tail = '') =>
    `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${MAIN}" xmlns:r="${DOC_RELS}"><sheetData>${rows}</sheetData>${tail}</worksheet>`;
  const column = (header: string, values: number[]) =>
    `<row r="1"><c r="A1" t="inlineStr"><is><t>${header}</t></is></c></row>` +
    values.map((value, at) => `<row r="${at + 2}"><c r="A${at + 2}"><v>${value}</v></c></row>`).join('');
  const table = (id: number, name: string, header: string) =>
    `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${MAIN}" id="${id}" name="${name}" displayName="${name}" ref="A1:A3"><tableColumns count="1"><tableColumn id="1" name="${header}"/></tableColumns></table>`;
  const tablePart = '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>';
  const override = (part: string, type: string) =>
    `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${type}+xml"/>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    override('/xl/workbook.xml', 'sheet.main') +
    [1, 2, 3, 4].map((n) => override(`/xl/worksheets/sheet${n}.xml`, 'worksheet')).join('') +
    [1, 2].map((n) => override(`/xl/tables/table${n}.xml`, 'table')).join('') + '</Types>');
  zip.file('_rels/.rels', rels([['rId1', 'officeDocument', 'xl/workbook.xml']]));
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${MAIN}" xmlns:r="${DOC_RELS}"><sheets>${['First', 'Data', 'Later', 'Other']
    .map((name, at) => `<sheet name="${name}" sheetId="${at + 1}" r:id="rId${at + 1}"/>`).join('')}</sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', rels([1, 2, 3, 4].map((n) => [`rId${n}`, 'worksheet', `worksheets/sheet${n}.xml`])));
  zip.file('xl/worksheets/sheet1.xml', sheet('<row r="1"><c r="A1"><v>100</v></c></row>'));
  zip.file('xl/worksheets/sheet2.xml', sheet(column('Amt', [1, 2]), tablePart));
  zip.file('xl/worksheets/_rels/sheet2.xml.rels', rels([['rIdT', 'table', '../tables/table1.xml']]));
  zip.file('xl/worksheets/sheet3.xml', sheet(column('Price', [10, 20]), tablePart));
  zip.file('xl/worksheets/_rels/sheet3.xml.rels', rels([['rIdT', 'table', '../tables/table2.xml']]));
  zip.file('xl/worksheets/sheet4.xml', sheet('<row r="1"><c r="A1"><f>SUM(Amounts[Amt])</f><v>3</v></c></row><row r="2"><c r="A2"><f>SUM(Prices[Price])</f><v>30</v></c></row>'));
  zip.file('xl/tables/table1.xml', table(1, 'Amounts', 'Amt'));
  zip.file('xl/tables/table2.xml', table(2, 'Prices', 'Price'));
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

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

  test('deleting and inserting sheets around tables, with undo and redo, matches on both engines', async () => {
    const bytes = await tablesFixture();
    const nativeWorkbook = native.openWorkbook(bytes, { calculation });
    const wasmWorkbook = wasm.openWorkbook(bytes, { calculation });
    const sums = (workbook: WorkbookHandle) => {
      const other = workbook.sheetInfo().sheetNames.length - 1;
      return [workbook.cell(other, 0, 0), workbook.cell(other, 1, 0)];
    };
    const script = (workbook: WorkbookHandle) => {
      let internal = 'applied';
      try {
        workbook.applyOps([{ type: 'setTables', tables: [] }]);
      } catch (error) {
        internal = String((error as Error)?.message ?? error);
      }
      const opened = sums(workbook);
      workbook.applyOps([{ type: 'removeSheet', index: 1 }]);
      const removed = sums(workbook);
      workbook.undo();
      const undone = sums(workbook);
      workbook.redo();
      const redone = sums(workbook);
      workbook.undo();
      workbook.applyOps([{ type: 'addSheet', index: 0, name: 'New' }]);
      const inserted = sums(workbook);
      return { internal, opened, removed, undone, redone, inserted };
    };
    try {
      const nativeResult = script(nativeWorkbook);
      const wasmResult = script(wasmWorkbook);
      expect(nativeResult).toEqual(wasmResult);
      expect(wasmResult.internal).toContain('restore sheet operations are internal');
      expect(wasmResult.undone).toEqual(wasmResult.opened);
      expect(wasmResult.inserted).toEqual(wasmResult.opened);
      expect(wasmResult.redone).toEqual(wasmResult.removed);
      await expectSavedPartsEqual(nativeWorkbook.save(), wasmWorkbook.save());
    } finally {
      nativeWorkbook.dispose();
      wasmWorkbook.dispose();
    }
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

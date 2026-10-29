import { readFileSync } from 'node:fs';

import {
  engineBackend,
  initWasm,
  openWorkbook,
  type EditResult,
  type WorkbookCalculationContext,
  type XlsxDefinedNamesRequest,
  type XlsxEditRequest,
  type XlsxExportOptions,
  type XlsxFindRequest,
  type XlsxMarkdownOptions,
  type XlsxReadRequest
} from '../../packages/xlsx/src/index';
import type { Format } from './harness';

function mutationResult({ applied, changed = [], limitedCells = [] }: EditResult) {
  return {
    applied,
    changed: [...changed].sort(),
    limitedCells: [...limitedCells].sort()
  };
}

export const xlsx: Format = {
  engineBackend,
  async open(bytes, script) {
    await initWasm(
      readFileSync(
        new URL('../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)
      )
    );
    const options = {
      calculation: script.calculation as WorkbookCalculationContext
    };
    let workbook = openWorkbook(bytes, options);
    return {
      operations: {
        open: (args) => {
          const next = openWorkbook(
            readFileSync(new URL(`../../${String(args.fixture)}`, import.meta.url)),
            options
          );
          workbook.dispose();
          workbook = next;
          return null;
        },
        sheetNames: () => workbook.sheetInfo().sheetNames,
        sheetInfo: () => workbook.sheetInfo(),
        listDefinedNames: (args) => workbook.listDefinedNames(args as XlsxDefinedNamesRequest),
        conditionalFormats: (args) => workbook.conditionalFormats(Number(args.sheet)),
        calculationStatus: () => workbook.calculationStatus(),
        cell: (args) => workbook.cell(Number(args.sheet), Number(args.row), Number(args.col)),
        version: () => workbook.version(),
        readCells: (args) => workbook.readCells(args as unknown as XlsxReadRequest),
        findText: (args) => workbook.findText(args as unknown as XlsxFindRequest),
        validateEdits: (args) => workbook.validateEdits(args as unknown as XlsxEditRequest),
        applyEdits: (args) => workbook.applyEdits(args as unknown as XlsxEditRequest),
        applyOps: (args) => mutationResult(workbook.applyOps(args.ops as unknown[])),
        historyState: () => workbook.historyState(),
        undo: () => mutationResult(workbook.undo()),
        redo: () => mutationResult(workbook.redo()),
        exportStructured: (args) => workbook.exportStructured(args as XlsxExportOptions),
        exportMarkdown: (args) => workbook.exportMarkdown(args.options as XlsxExportOptions, args.markdownOptions as XlsxMarkdownOptions),
        save: () => workbook.save()
      },
      dispose: () => workbook.dispose()
    };
  }
};

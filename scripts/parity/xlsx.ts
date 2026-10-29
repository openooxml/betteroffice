import { readFileSync } from 'node:fs';

import {
  engineBackend,
  initWasm,
  openWorkbook,
  type WorkbookCalculationContext,
  type XlsxEditRequest,
  type XlsxExportOptions,
  type XlsxFindRequest,
  type XlsxMarkdownOptions,
  type XlsxReadRequest
} from '../../packages/xlsx/src/index';
import type { Format } from './harness';

export const xlsx: Format = {
  engineBackend,
  async open(bytes, script) {
    await initWasm(
      readFileSync(
        new URL('../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)
      )
    );
    const workbook = openWorkbook(bytes, {
      calculation: script.calculation as WorkbookCalculationContext
    });
    return {
      operations: {
        sheetNames: () => workbook.sheetInfo().sheetNames,
        version: () => workbook.version(),
        readCells: (args) => workbook.readCells(args as unknown as XlsxReadRequest),
        findText: (args) => workbook.findText(args as unknown as XlsxFindRequest),
        validateEdits: (args) => workbook.validateEdits(args as unknown as XlsxEditRequest),
        applyEdits: (args) => workbook.applyEdits(args as unknown as XlsxEditRequest),
        historyState: () => workbook.historyState(),
        undo: () => workbook.undo().applied,
        redo: () => workbook.redo().applied,
        exportStructured: (args) => workbook.exportStructured(args as XlsxExportOptions),
        exportMarkdown: (args) => workbook.exportMarkdown(args.options as XlsxExportOptions, args.markdownOptions as XlsxMarkdownOptions),
        save: () => workbook.save()
      },
      dispose: () => workbook.dispose()
    };
  }
};

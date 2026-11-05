import { readFileSync } from 'node:fs';

import {
  initWasm,
  openWorkbook,
  type CapturedFormat,
  type ChartMutation,
  type EditResult,
  type MergedRange,
  type WorkbookCalculationContext,
  type WorkbookHyperlink,
  type XlsxDefinedNamesRequest,
  type XlsxEditRequest,
  type XlsxExportOptions,
  type XlsxFindRequest,
  type XlsxMarkdownOptions,
  type XlsxReadRequest
} from '../../packages/xlsx/src/index';
import { engineKind } from '../../packages/xlsx/src/wasm/engine';
import type { Format } from './harness';

function mutationResult({ applied, changed = [], limitedCells = [] }: EditResult) {
  return {
    applied,
    changed: [...changed].sort(),
    limitedCells: [...limitedCells].sort()
  };
}

export const xlsx: Format = {
  engineKind,
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
        openCollaborative: (args) => {
          const next = openWorkbook(
            readFileSync(new URL(`../../${String(args.fixture)}`, import.meta.url)),
            { collaborative: true, clientId: Number(args.clientId) }
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
        rangeCells: (args) => workbook.rangeCells(Number(args.sheet), String(args.range)),
        searchText: (args) => workbook.searchText(String(args.query), args.options as { caseSensitive?: boolean; limit?: number }),
        selectionFormatting: (args) => workbook.selectionFormatting(Number(args.sheet), String(args.range)),
        charts: (args) => workbook.charts(Number(args.sheet)),
        version: () => workbook.version(),
        readCells: (args) => workbook.readCells(args as unknown as XlsxReadRequest),
        findText: (args) => workbook.findText(args as unknown as XlsxFindRequest),
        validateEdits: (args) => workbook.validateEdits(args as unknown as XlsxEditRequest),
        applyEdits: (args) => workbook.applyEdits(args as unknown as XlsxEditRequest),
        applyOps: (args) => mutationResult(workbook.applyOps(args.ops as unknown[])),
        removeNewestConditionalFormat: (args) => {
          const sheet = Number(args.sheet);
          const rules = workbook.conditionalFormats(sheet);
          const { id } = rules.reduce((newest, rule) => rule.priority > newest.priority ? rule : newest);
          return mutationResult(workbook.applyOps([{ type: 'removeConditionalFormat', sheet, id }]));
        },
        moveRange: (args) => mutationResult(
          workbook.moveRange(Number(args.sheet), String(args.source), String(args.destination))
        ),
        setHyperlink: (args) => mutationResult(
          workbook.setHyperlink(Number(args.sheet), args.hyperlink as WorkbookHyperlink)
        ),
        removeHyperlink: (args) => mutationResult(
          workbook.removeHyperlink(Number(args.sheet), args.range as MergedRange)
        ),
        editCharts: (args) => mutationResult(workbook.editCharts(args.edits as ChartMutation[])),
        moveChart: (args) => mutationResult(
          workbook.moveChart(Number(args.sheet), String(args.chart), Number(args.dx), Number(args.dy))
        ),
        captureFormat: (args) => workbook.captureFormat(Number(args.sheet), String(args.range)),
        applyFormat: (args) => mutationResult(
          workbook.applyFormat(Number(args.sheet), String(args.range), args.format as CapturedFormat)
        ),
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

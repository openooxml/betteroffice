import type { CellAddr, CellInputEdit, WorkbookHandle } from '@betteroffice/xlsx';
import { XlsxHostRefusalError } from '../commands/createXlsxCommandStore';
import type { XlsxPluginHost } from './createXlsxPluginHost';
import type { XlsxPluginCellCommit, XlsxPluginCellRef } from './types';

type GateWorkbook = Pick<WorkbookHandle, 'readCells'>;

/** `cells` of `sheet` as plugins gate them, or null when no plugin gates edits. */
export function gatedCells(
  gate: XlsxPluginHost, workbook: GateWorkbook, sheet: number, cells: readonly CellAddr[]
): XlsxPluginCellRef[] | null {
  if (cells.length === 0 || !gate.gatesEdits()) return null;
  const catalog = workbook.readCells({ ranges: [] });
  const entry = catalog.ok ? catalog.sheets[sheet] : undefined;
  return entry ? cells.map(({ row, col }) => ({ sheetId: entry.sheetId, row, col })) : null;
}

/** Whether plugins let the user type into a cell. */
export function cellEditable(gate: XlsxPluginHost, workbook: GateWorkbook, sheet: number, cell: CellAddr): boolean {
  const [ref] = gatedCells(gate, workbook, sheet, [cell]) ?? [];
  return !ref || gate.openGate().cellEditable(ref);
}

/**
 * Runs typed writes through the plugins: the inputs to write, or the refusal, whose message is
 * the plugin's or `fallback`.
 */
export function gateWrites(
  gate: XlsxPluginHost, workbook: GateWorkbook, sheet: number, edits: readonly CellInputEdit[],
  source: XlsxPluginCellCommit['source'], fallback: string
): CellInputEdit[] | XlsxHostRefusalError {
  const cells = gatedCells(gate, workbook, sheet, edits);
  if (!cells) return [...edits];
  const pass = gate.openGate();
  if (!cells.every((cell) => pass.cellEditable(cell))) return new XlsxHostRefusalError(fallback);
  const gated: CellInputEdit[] = [];
  for (const [index, cell] of cells.entries()) {
    const outcome = pass.beforeCellCommit({ ...cell, input: edits[index].input, source });
    if (!outcome.ok) return new XlsxHostRefusalError(outcome.message ?? fallback);
    gated.push({ ...edits[index], input: outcome.input });
  }
  return gated;
}

/**
 * A clear through the plugins: only cells holding content are asked, and the result writes the
 * inputs a plugin replaced over the cleared cells.
 */
export function gateClear(
  gate: XlsxPluginHost, workbook: GateWorkbook & Pick<WorkbookHandle, 'cell'>, sheet: number,
  edits: readonly CellInputEdit[], fallback: string
): CellInputEdit[] | XlsxHostRefusalError {
  if (!gate.gatesEdits()) return [...edits];
  const content = edits.filter(({ row, col }) => workbook.cell(sheet, row, col).input !== '');
  const gated = gateWrites(gate, workbook, sheet, content, 'clear-key', fallback);
  if (!Array.isArray(gated)) return gated;
  const inputs = new Map(gated.map((edit) => [`${edit.row}:${edit.col}`, edit.input]));
  return edits.map((edit) => ({ ...edit, input: inputs.get(`${edit.row}:${edit.col}`) ?? edit.input }));
}

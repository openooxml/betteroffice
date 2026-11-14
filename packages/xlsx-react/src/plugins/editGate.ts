import { translateFormulasOn } from '@betteroffice/xlsx';
import type { CellAddr, CellInputEdit, CellPoint, CellRange, WorkbookHandle } from '@betteroffice/xlsx';
import { XlsxHostRefusalError } from '../commands/createXlsxCommandStore';
import type { XlsxPluginHost } from './createXlsxPluginHost';
import type { XlsxPluginCellCommit, XlsxPluginCellRef } from './types';

/** The most cells one paste or cut asks plugins about; a larger one is refused. */
export const GATE_CELL_LIMIT = 50_000;
const MAX_INPUT_READ_CELLS = 100_000;

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
 * the plugin's or `fallback`, or `starting` while a gating plugin has not loaded yet.
 */
export function gateWrites(
  gate: XlsxPluginHost, workbook: GateWorkbook, sheet: number, edits: readonly CellInputEdit[],
  source: XlsxPluginCellCommit['source'], fallback: string, starting: string
): CellInputEdit[] | XlsxHostRefusalError {
  if (edits.length > 0 && gate.editGateStarting()) return new XlsxHostRefusalError(starting);
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
 * A clear or cut through the plugins: only cells holding content are asked, and the result writes
 * the inputs a plugin replaced over the cleared cells.
 */
export function gateClear(
  gate: XlsxPluginHost, workbook: GateWorkbook & Pick<WorkbookHandle, 'cell'>, sheet: number,
  edits: readonly CellInputEdit[], fallback: string, starting: string, cut?: { tooMany: string }
): CellInputEdit[] | XlsxHostRefusalError {
  if (!gate.gatesEdits()) return [...edits];
  const content = edits.filter(({ row, col }) => workbook.cell(sheet, row, col).input !== '');
  if (cut && content.length > GATE_CELL_LIMIT) return new XlsxHostRefusalError(cut.tooMany);
  const source = cut ? 'cut' : 'clear-key';
  const gated = gateWrites(gate, workbook, sheet, content, source, fallback, starting);
  if (!Array.isArray(gated)) return gated;
  const inputs = new Map(gated.map((edit) => [`${edit.row}:${edit.col}`, edit.input]));
  return edits.map((edit) => ({ ...edit, input: inputs.get(`${edit.row}:${edit.col}`) ?? edit.input }));
}

/** 0 is A, 26 is AA. */
function columnLetters(col: number): string {
  let rest = col + 1;
  let letters = '';
  while (rest > 0) {
    rest -= 1;
    letters = String.fromCharCode(65 + (rest % 26)) + letters;
    rest = Math.floor(rest / 26);
  }
  return letters;
}

export function rangeA1(range: CellRange): string {
  return `${columnLetters(range.left)}${range.top + 1}:${columnLetters(range.right)}${range.bottom + 1}`;
}

type FormulaMoves = { at: number[]; items: { formula: string; from: CellPoint; to: CellPoint }[] };
type MovedFormula = { ok: true; formula: string } | { ok: false; error: string };

/** The formulas among `edits` copied inside the editor, which the engine moves to their cells. */
export function formulaMoves(edits: readonly CellInputEdit[]): FormulaMoves {
  const moves: FormulaMoves = { at: [], items: [] };
  edits.forEach((edit, at) => {
    if (!edit.copyFrom || !edit.input.startsWith('=')) return;
    moves.at.push(at);
    moves.items.push({ formula: edit.input.slice(1), from: edit.copyFrom, to: { row: edit.row, col: edit.col } });
  });
  return moves;
}

/** The input the engine writes for each edit, with the moved formulas in the order of `moves`. */
export function writtenInputs(edits: readonly CellInputEdit[], moves: FormulaMoves, moved: readonly MovedFormula[]): string[] {
  const written = edits.map((edit) => edit.input);
  moves.at.forEach((at, index) => {
    const result = moved[index];
    if (!result?.ok) throw new Error(result?.error ?? 'formula could not be moved');
    written[at] = `=${result.formula}`;
  });
  return written;
}

/** The ranges to read for the inputs under `edits`: their bounds, or each row of them when large. */
export function inputRanges(edits: readonly CellInputEdit[]): CellRange[] {
  if (edits.length === 0) return [];
  const rows = new Map<number, CellRange>();
  for (const { row, col } of edits) {
    const range = rows.get(row);
    if (!range) rows.set(row, { top: row, bottom: row, left: col, right: col });
    else {
      range.left = Math.min(range.left, col);
      range.right = Math.max(range.right, col);
    }
  }
  const lines = [...rows.values()];
  const bounds = lines.reduce((all, line) => ({
    top: Math.min(all.top, line.top), bottom: Math.max(all.bottom, line.bottom),
    left: Math.min(all.left, line.left), right: Math.max(all.right, line.right),
  }));
  const area = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  return area <= MAX_INPUT_READ_CELLS ? [bounds] : lines;
}

/**
 * The edits that change their cell, each with the input the engine writes there; `read` holds the
 * inputs of `ranges` before the paste, in their order.
 */
export function changedEdits(
  edits: readonly CellInputEdit[], written: readonly string[], ranges: readonly CellRange[],
  read: readonly (readonly { input: string }[])[][]
): CellInputEdit[] {
  const current = (row: number, col: number) => {
    const at = ranges.findIndex((range) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);
    return at < 0 ? '' : read[at][row - ranges[at].top]?.[col - ranges[at].left]?.input ?? '';
  };
  const lookup = ranges.length === 1 ? current : (() => {
    const byRow = new Map(ranges.map((range, at) => [range.top, at]));
    return (row: number, col: number) => {
      const at = byRow.get(row);
      return at === undefined ? '' : read[at][0]?.[col - ranges[at].left]?.input ?? '';
    };
  })();
  const changed: CellInputEdit[] = [];
  edits.forEach((edit, at) => {
    const input = written[at];
    if (input !== lookup(edit.row, edit.col)) changed.push({ row: edit.row, col: edit.col, input });
  });
  return changed;
}

/** `edits` with every input a plugin replaced, written as typed instead of copied. */
export function withGatedInputs(
  edits: readonly CellInputEdit[], changed: readonly CellInputEdit[], gated: readonly CellInputEdit[]
): CellInputEdit[] {
  const replaced = new Map<string, string>();
  gated.forEach((edit, at) => { if (edit.input !== changed[at].input) replaced.set(`${edit.row}:${edit.col}`, edit.input); });
  if (replaced.size === 0) return [...edits];
  return edits.map((edit) => {
    const input = replaced.get(`${edit.row}:${edit.col}`);
    return input === undefined ? edit : { row: edit.row, col: edit.col, input };
  });
}

/**
 * A paste through the plugins: each cell it changes is asked once, with the input written there,
 * and any refusal refuses the whole paste.
 */
export function gatePaste(
  gate: XlsxPluginHost, workbook: WorkbookHandle, sheet: number,
  edits: readonly CellInputEdit[], fallback: string, starting: string, tooMany: string
): CellInputEdit[] | XlsxHostRefusalError {
  if (edits.length === 0 || !gate.gatesEdits()) return [...edits];
  const moves = formulaMoves(edits);
  const written = writtenInputs(edits, moves, moves.items.length ? translateFormulasOn(workbook, moves.items) : []);
  const ranges = inputRanges(edits);
  const changed = changedEdits(edits, written, ranges, ranges.map((range) => workbook.rangeCells(sheet, rangeA1(range))));
  if (changed.length > GATE_CELL_LIMIT) return new XlsxHostRefusalError(tooMany);
  const gated = gateWrites(gate, workbook, sheet, changed, 'paste', fallback, starting);
  return Array.isArray(gated) ? withGatedInputs(edits, changed, gated) : gated;
}

/** Rectangles covering exactly the cells of `edits`: one when they fill their bounds. */
export function coverRanges(edits: readonly CellAddr[]): CellRange[] {
  const runs: CellRange[] = [];
  for (const { row, col } of [...edits].sort((a, b) => a.row - b.row || a.col - b.col)) {
    const last = runs[runs.length - 1];
    if (last && last.top === row && col <= last.right + 1) last.right = Math.max(last.right, col);
    else runs.push({ top: row, bottom: row, left: col, right: col });
  }
  const covered: CellRange[] = [];
  for (const run of runs) {
    const last = covered[covered.length - 1];
    if (last && last.left === run.left && last.right === run.right && last.bottom + 1 === run.top) last.bottom = run.top;
    else covered.push(run);
  }
  return covered;
}

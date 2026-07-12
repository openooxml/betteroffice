/**
 * A tiny demo "agent" for the example app: it inspects the active sheet through
 * the workbook handle, finds numeric columns, and builds `=SUM(...)` proposals
 * one row below the data. Staging these via `handle.propose` exercises the full
 * human-in-the-loop review flow with a real, useful edit. Example-only — not
 * part of the shipped library.
 */

import type { CellEdit, ProposalEdit, WorkbookHandle } from '@betteroffice/xlsx';

// scan a bounded window; enough to cover the sample without walking a huge sheet.
const SCAN_ROWS = 200;
const SCAN_COLS = 26;
// a column needs at least this many numbers to earn a total (skips lone header
// digits like a year).
const MIN_NUMERIC = 2;

// bijective base-26 column letter: 0 -> A, 25 -> Z, 26 -> AA.
function columnLetter(col: number): string {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

// a literal number cell — not a formula, not text. the editable input is the
// source of truth; computed values aren't visible from CellEdit.
function isNumericLiteral(cell: CellEdit): boolean {
  const text = cell.input.trim();
  if (text === '' || cell.isFormula) return false;
  return Number.isFinite(Number(text));
}

interface ColumnSpan {
  first: number;
  last: number;
  count: number;
}

/**
 * Build `=SUM` total edits for every numeric column in the active sheet, placed
 * on the first empty row below the data. Empty when nothing numeric is found.
 */
export function buildTotalsEdits(handle: WorkbookHandle): ProposalEdit[] {
  const sheet = handle.sheetInfo().activeSheet;
  let rows: CellEdit[][];
  try {
    rows = handle.rangeCells(sheet, `A1:${columnLetter(SCAN_COLS - 1)}${SCAN_ROWS}`);
  } catch {
    return [];
  }

  const columns = new Map<number, ColumnSpan>();
  let lastDataRow = -1;
  rows.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (!isNumericLiteral(cell)) return;
      const span = columns.get(c) ?? { first: r, last: r, count: 0 };
      span.first = Math.min(span.first, r);
      span.last = Math.max(span.last, r);
      span.count += 1;
      columns.set(c, span);
      lastDataRow = Math.max(lastDataRow, r);
    });
  });
  if (lastDataRow < 0) return [];

  const totalRow = lastDataRow + 1;
  const edits: ProposalEdit[] = [];
  for (const [col, span] of columns) {
    if (span.count < MIN_NUMERIC) continue;
    const letter = columnLetter(col);
    edits.push({
      sheet,
      row: totalRow,
      col,
      input: `=SUM(${letter}${span.first + 1}:${letter}${span.last + 1})`,
    });
  }
  return edits;
}

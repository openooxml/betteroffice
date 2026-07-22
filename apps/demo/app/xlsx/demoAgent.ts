import type { CellEdit, ProposalEdit, WorkbookHandle } from "@betteroffice/xlsx";

const SCAN_ROWS = 200;
const SCAN_COLS = 26;
const MIN_NUMERIC = 2;

function columnLetter(col: number): string {
  let n = col;
  let out = "";
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

function isNumericLiteral(cell: CellEdit): boolean {
  const text = cell.input.trim();
  return text !== "" && !cell.isFormula && Number.isFinite(Number(text));
}

interface ColumnSpan {
  first: number;
  last: number;
  count: number;
}

function dashboardEdits(rows: CellEdit[][], sheet: number): ProposalEdit[] | null {
  const headerRow = rows.findIndex((row) =>
    ["Units", "Revenue", "Growth", "Updated"].every((label) =>
      row.some((cell) => cell.input.trim() === label),
    ),
  );
  if (headerRow < 0) return null;

  const column = (label: string) =>
    rows[headerRow].findIndex((cell) => cell.input.trim() === label);
  const units = column("Units");
  const revenue = column("Revenue");
  const growth = column("Growth");
  const updated = column("Updated");
  const totalRow = rows.findIndex(
    (row, index) => index > headerRow && row[0]?.input.trim() === "Total",
  );
  if ([units, revenue, growth, updated, totalRow].some((index) => index < 0)) return null;

  const dataRows = rows
    .map((row, index) => ({ row, index }))
    .filter(({ row, index }) => index > headerRow && index < totalRow && row[0]?.input.trim());
  const unitsCell = dataRows.find(({ row }) => isNumericLiteral(row[units]));
  const revenueCell = dataRows.find(({ row }) => isNumericLiteral(row[revenue]));
  if (!unitsCell || !revenueCell) return null;

  const edits: ProposalEdit[] = [
    {
      sheet,
      row: unitsCell.index,
      col: units,
      input: String(Number(unitsCell.row[units].input) + 60),
    },
    {
      sheet,
      row: revenueCell.index,
      col: revenue,
      input: String(Number(revenueCell.row[revenue].input) + 5000),
    },
  ];
  const firstDataRow = dataRows[0].index + 1;
  const lastDataRow = dataRows[dataRows.length - 1].index + 1;
  if (rows[totalRow][growth].input.trim() === "") {
    const letter = columnLetter(growth);
    edits.push({
      sheet,
      row: totalRow,
      col: growth,
      input: `=SUM(${letter}${firstDataRow}:${letter}${lastDataRow})`,
      numberFormat: "percent",
    });
  }
  if (rows[totalRow][updated].input.trim() === "") {
    const letter = columnLetter(updated);
    edits.push({
      sheet,
      row: totalRow,
      col: updated,
      input: `=MAX(${letter}${firstDataRow}:${letter}${lastDataRow})`,
      numberFormat: "date",
    });
  }
  return edits.filter((edit) => rows[edit.row][edit.col].input !== edit.input);
}

function genericTotals(rows: CellEdit[][], sheet: number): ProposalEdit[] {
  const columns = new Map<number, ColumnSpan>();
  let lastDataRow = -1;
  rows.forEach((row, rowIndex) => {
    row.forEach((cell, col) => {
      if (!isNumericLiteral(cell)) return;
      const span = columns.get(col) ?? { first: rowIndex, last: rowIndex, count: 0 };
      span.first = Math.min(span.first, rowIndex);
      span.last = Math.max(span.last, rowIndex);
      span.count += 1;
      columns.set(col, span);
      lastDataRow = Math.max(lastDataRow, rowIndex);
    });
  });
  if (lastDataRow < 0) return [];

  const totalRow = lastDataRow + 1;
  const edits: ProposalEdit[] = [];
  for (const [col, span] of columns) {
    if (span.count < MIN_NUMERIC) continue;
    const letter = columnLetter(col);
    const input = `=SUM(${letter}${span.first + 1}:${letter}${span.last + 1})`;
    if (rows[totalRow]?.[col]?.input === input) continue;
    edits.push({ sheet, row: totalRow, col, input });
  }
  return edits;
}

/** Build the demo's tracked totals proposal. */
export function buildTotalsEdits(handle: WorkbookHandle): ProposalEdit[] {
  const sheet = handle.sheetInfo().activeSheet;
  let rows: CellEdit[][];
  try {
    rows = handle.rangeCells(sheet, `A1:${columnLetter(SCAN_COLS - 1)}${SCAN_ROWS}`);
  } catch {
    return [];
  }
  return dashboardEdits(rows, sheet) ?? genericTotals(rows, sheet);
}

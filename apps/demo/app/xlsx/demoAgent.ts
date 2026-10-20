import type {
  ProposalEdit,
  XlsxExportCell,
  XlsxExportSheet,
  XlsxStructuredContent,
} from "@betteroffice/xlsx";

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

interface GridCell {
  input: string;
  isFormula: boolean;
  isNumber: boolean;
  occupied: boolean;
}

function cellPosition(a1: string): { row: number; col: number } | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/.exec(a1);
  if (!match) return null;
  let col = 0;
  for (const letter of match[1]) col = col * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, col: col - 1 };
}

function cellInput(cell: XlsxExportCell): GridCell {
  if (cell.formula !== null) return { input: `=${cell.formula}`, isFormula: true, isNumber: false, occupied: true };
  const value = cell.value;
  const input =
    value.kind === "number" || value.kind === "text"
      ? String(value.value)
      : value.kind === "bool"
        ? (value.value ? "TRUE" : "FALSE")
        : "";
  return { input, isFormula: false, isNumber: value.kind === "number", occupied: value.kind !== "empty" };
}

function sheetGrid(sheet: XlsxExportSheet): GridCell[][] {
  const rows = Array.from({ length: SCAN_ROWS }, () =>
    Array.from({ length: SCAN_COLS }, () => ({ input: "", isFormula: false, isNumber: false, occupied: false })),
  );
  for (const cell of sheet.cells) {
    if (cell.anchor.kind !== "cell") continue;
    const at = cellPosition(cell.anchor.a1);
    if (at && at.row < SCAN_ROWS && at.col < SCAN_COLS) {
      rows[at.row][at.col] = cellInput(cell);
    }
  }
  return rows;
}

interface ColumnSpan {
  first: number;
  last: number;
  count: number;
}

function dashboardEdits(rows: GridCell[][], sheet: number): ProposalEdit[] | null {
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
  const unitsCell = dataRows.find(({ row }) => row[units].isNumber);
  const revenueCell = dataRows.find(({ row }) => row[revenue].isNumber);
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
  if (!rows[totalRow][growth].occupied) {
    const letter = columnLetter(growth);
    edits.push({
      sheet,
      row: totalRow,
      col: growth,
      input: `=SUM(${letter}${firstDataRow}:${letter}${lastDataRow})`,
      numberFormat: "percent",
    });
  }
  if (!rows[totalRow][updated].occupied) {
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

function genericTotals(rows: GridCell[][], sheet: number): ProposalEdit[] {
  const columns = new Map<number, ColumnSpan>();
  let lastDataRow = -1;
  rows.forEach((row, rowIndex) => {
    row.forEach((cell, col) => {
      if (!cell.isNumber) return;
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
    if (totalRow >= SCAN_ROWS || rows[totalRow][col].occupied) continue;
    edits.push({ sheet, row: totalRow, col, input });
  }
  return edits;
}

/** Build the demo's tracked totals proposal from the first worksheet that yields one. */
export function buildTotalsEdits(content: XlsxStructuredContent): ProposalEdit[] {
  const sheets = content.sheets.flatMap((sheet, position) => {
    if (sheet.kind !== "worksheet" || sheet.truncated) return [];
    const index = sheet.anchor.kind === "sheet" ? sheet.anchor.sheet.index : position;
    return [{ rows: sheetGrid(sheet), index }];
  });
  for (const { rows, index } of sheets) {
    const edits = dashboardEdits(rows, index);
    if (edits?.length) return edits;
  }
  for (const { rows, index } of sheets) {
    const edits = genericTotals(rows, index);
    if (edits.length) return edits;
  }
  return [];
}

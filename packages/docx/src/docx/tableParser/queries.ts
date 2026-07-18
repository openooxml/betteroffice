/**
 * Table Query Helpers
 *
 * Read-only utilities for inspecting a parsed Table model — column/row
 * counts, cell merge state, plain-text extraction, header detection,
 * floating-table detection. No XML access; these operate on the parsed
 * `Table` shape.
 */

import type { Table, TableRow, TableCell, Paragraph } from '../../types/document';

/**
 * Get the number of columns in a table
 *
 * Uses the table grid if available, otherwise counts cells in first row.
 *
 * @param table - The table to measure
 * @returns Number of columns
 */
export function getTableColumnCount(table: Table): number {
  if (table.columnWidths && table.columnWidths.length > 0) {
    return table.columnWidths.length;
  }

  if (table.rows.length === 0) return 0;

  // Count cells in first row, accounting for grid span
  return table.rows[0].cells.reduce((count, cell) => {
    return count + (cell.formatting?.gridSpan ?? 1);
  }, 0);
}

/**
 * Get the number of rows in a table
 *
 * @param table - The table to measure
 * @returns Number of rows
 */
export function getTableRowCount(table: Table): number {
  return table.rows.length;
}

/**
 * Check if a cell is part of a vertical merge
 *
 * @param cell - The cell to check
 * @returns true if cell continues a vertical merge
 */
export function isCellMergeContinuation(cell: TableCell): boolean {
  return cell.formatting?.vMerge === 'continue';
}

/**
 * Check if a cell starts a vertical merge
 *
 * @param cell - The cell to check
 * @returns true if cell starts a vertical merge
 */
export function isCellMergeStart(cell: TableCell): boolean {
  return cell.formatting?.vMerge === 'restart';
}

/**
 * Check if a cell spans multiple columns
 *
 * @param cell - The cell to check
 * @returns true if cell spans multiple columns
 */
export function isCellHorizontallyMerged(cell: TableCell): boolean {
  return (cell.formatting?.gridSpan ?? 1) > 1;
}

/**
 * Get the plain text content of a table
 *
 * @param table - The table to extract text from
 * @returns Plain text content
 */
export function getTableText(table: Table): string {
  const rows: string[] = [];

  for (const row of table.rows) {
    const cells: string[] = [];

    for (const cell of row.cells) {
      const cellText = cell.content
        .filter((c): c is Paragraph => c.type === 'paragraph')
        .map((p) => getParagraphText(p))
        .join('\n');
      cells.push(cellText);
    }

    rows.push(cells.join('\t'));
  }

  return rows.join('\n');
}

/**
 * Helper to get paragraph text (simplified)
 */
function getParagraphText(para: Paragraph): string {
  return para.content
    .filter((c) => 'content' in c)
    .flatMap((run) => {
      if (!('content' in run)) return [];
      return (run as { content: Array<{ type: string; text?: string }> }).content
        .filter((c) => c.type === 'text')
        .map((c) => c.text ?? '');
    })
    .join('');
}

/**
 * Check if table has header row
 *
 * @param table - The table to check
 * @returns true if first row is marked as header
 */
export function hasHeaderRow(table: Table): boolean {
  if (table.rows.length === 0) return false;
  return table.rows[0].formatting?.header === true;
}

/**
 * Get all header rows from a table
 *
 * @param table - The table to search
 * @returns Array of header rows
 */
export function getHeaderRows(table: Table): TableRow[] {
  return table.rows.filter((row) => row.formatting?.header === true);
}

/**
 * Check if table is a floating table
 *
 * @param table - The table to check
 * @returns true if table has floating properties
 */
export function isFloatingTable(table: Table): boolean {
  return table.formatting?.floating !== undefined;
}

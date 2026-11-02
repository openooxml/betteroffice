import type { AwarenessCursor } from '@betteroffice/xlsx/collaboration';

export const AWARENESS_LABEL_DURATION_MS = 3_000;

const XLSX_MAX_ROWS = 1_048_576;
const XLSX_MAX_COLUMNS = 16_384;

const AWARENESS_COLORS = [
  '#0B57D0',
  '#B3261E',
  '#146C2E',
  '#6B4EFF',
  '#8E24AA',
  '#00639B',
  '#8E4E00',
  '#7D5260',
] as const;

export function normalizeAwarenessColor(color: unknown, clientId: number): string {
  if (typeof color !== 'string' || !/^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(color)) {
    return AWARENESS_COLORS[clientId % AWARENESS_COLORS.length];
  }
  const hex = color.slice(1).toUpperCase();
  if (hex.length === 6) return `#${hex}`;
  return `#${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`;
}

export function resolveAwarenessCursor(
  cursor: AwarenessCursor,
  sheetIds: readonly string[],
  rows = XLSX_MAX_ROWS,
  columns = XLSX_MAX_COLUMNS
): { sheetIndex: number; range: { top: number; left: number; bottom: number; right: number } } | null {
  const sheetIndex = sheetIds.indexOf(cursor.sheet);
  if (sheetIndex < 0 || rows < 1 || columns < 1) return null;
  const cells = [cursor.anchor, cursor.head];
  if (
    cells.some(
      ({ row, col }) =>
        !Number.isSafeInteger(row) ||
        !Number.isSafeInteger(col) ||
        row < 0 ||
        col < 0 ||
        row >= rows ||
        col >= columns
    )
  ) {
    return null;
  }
  return {
    sheetIndex,
    range: {
      top: Math.min(cursor.anchor.row, cursor.head.row),
      left: Math.min(cursor.anchor.col, cursor.head.col),
      bottom: Math.max(cursor.anchor.row, cursor.head.row),
      right: Math.max(cursor.anchor.col, cursor.head.col),
    },
  };
}

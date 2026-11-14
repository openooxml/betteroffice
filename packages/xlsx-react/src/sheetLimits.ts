/**
 * The sheet's last addressable row and column. Kept beside the grid components
 * so an eager surface can size a whole-sheet selection without importing the
 * collaboration runtime (the constants there are re-exported from awareness).
 * Mirrors `XLSX_MAX_ROWS`/`XLSX_MAX_COLUMNS` in `@betteroffice/xlsx/collaboration`.
 */
export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLUMNS = 16_384;

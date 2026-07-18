/**
 * Shared table-width helpers used by both the agent table-insert command and
 * the DOCX table-grid serializer, so the default content width and the
 * even-split algorithm have a single source of truth.
 */

/** Default table content width in twips (≈ 6.5in usable width on Letter). */
export const DEFAULT_TABLE_WIDTH_DXA = 9360;

/**
 * Split `totalWidth` into `columnCount` integer column widths that sum to
 * `totalWidth`, dripping the rounding remainder across the leading columns.
 * Every returned width is at least 1.
 */
export function distributeColumnWidths(totalWidth: number, columnCount: number): number[] {
  if (columnCount <= 0) return [];

  const baseWidth = Math.floor(totalWidth / columnCount);
  let remainder = totalWidth - baseWidth * columnCount;

  return Array.from({ length: columnCount }, () => {
    const width = baseWidth + (remainder > 0 ? 1 : 0);
    remainder = Math.max(0, remainder - 1);
    return Math.max(1, width);
  });
}

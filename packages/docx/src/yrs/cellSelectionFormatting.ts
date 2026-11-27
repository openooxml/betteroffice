import type { YrsSession, YrsStoryRange, YrsTableRange } from './index';
import { tableAnchors, type TablePayload } from './storyPlainText';

/** A run mark the selection context reports as on, off or mixed. @internal */
export type YrsToggledMark = 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript';

/** The whole text of each non-empty cell that `cells` covers, in reading order. @internal */
export function cellSelectionRanges(
  reader: Pick<YrsSession, 'tablePayload' | 'paragraphSpans'>,
  cells: YrsTableRange
): YrsStoryRange[] {
  const payload = reader.tablePayload(cells.anchor.story, cells.anchor.tableIndex);
  if (!payload || !Array.isArray(payload.rows)) return [];
  const top = Math.min(cells.anchor.row, cells.head.row);
  const bottom = Math.max(cells.anchor.row, cells.head.row);
  const left = Math.min(cells.anchor.column, cells.head.column);
  const right = Math.max(cells.anchor.column, cells.head.column);
  return tableAnchors(payload as unknown as TablePayload).anchors.flatMap((cell) => {
    if (!cell.story || cell.row > bottom || cell.row + cell.rowspan <= top ||
      cell.column > right || cell.column + cell.colspan <= left) return [];
    const spans = reader.paragraphSpans(cell.story);
    const first = spans[0];
    const last = spans[spans.length - 1];
    if (!first || !last || (first === last && last.length === 0)) return [];
    return [{ story: cell.story, start: { paraId: first.paraId, offset: 0 }, end: { paraId: last.paraId, offset: last.length } }];
  });
}

/**
 * Toggles `mark` over `ranges` as one selection: off where it is set once all their text carries
 * it, otherwise on where it is not. A range whose story holds no text, such as an empty cell,
 * does not keep it from turning off.
 * @internal
 */
export function toggleAcross(
  session: Pick<YrsSession, 'paragraphs' | 'selectionContext'>,
  ranges: readonly YrsStoryRange[],
  mark: YrsToggledMark,
  toggle: (range: YrsStoryRange, on: boolean) => void
): void {
  const carried = ranges.map((range) => session.selectionContext(range)[mark] === true);
  const off = carried.some(Boolean) && ranges.every((range, index) =>
    carried[index] || session.paragraphs(range.story).every((paragraph) => paragraph.text === ''));
  ranges.forEach((range, index) => {
    if (carried[index] === off) toggle(range, !off);
  });
}

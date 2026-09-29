import { renderedScale } from '@betteroffice/docx/layout/render';

/**
 * The band of `column` the viewport shows, in the column's layout pixels, the
 * space its page offsets are in. `scroller` is null for the root scroller.
 */
export function viewportColumnBand(
  scroller: HTMLElement | null,
  column: HTMLElement
): { top: number; bottom: number } {
  const scrollerRect = scroller?.getBoundingClientRect() ?? null;
  const viewportTop = scrollerRect?.top ?? 0;
  const viewportHeight =
    scroller && scrollerRect
      ? scroller.clientHeight * renderedScale(scroller, scrollerRect)
      : window.innerHeight;
  const columnRect = column.getBoundingClientRect();
  const scale = renderedScale(column, columnRect);
  const top = (viewportTop - columnRect.top) / scale;
  return { top, bottom: top + viewportHeight / scale };
}

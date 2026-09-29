import { effectiveZoom } from '@betteroffice/docx/layout/render';

/**
 * The client-pixel band a scroller shows, and the client pixels one unit of its
 * `scrollTop` moves. The root scrolls in viewport units whatever its zoom.
 */
export interface ScrollViewport {
  top: number;
  bottom: number;
  zoom: number;
}

function isRootScroller(scroller: Element): boolean {
  return (
    typeof document !== 'undefined' &&
    (scroller === document.scrollingElement ||
      scroller === document.documentElement ||
      scroller === document.body)
  );
}

export function scrollViewport(scroller: HTMLElement): ScrollViewport {
  if (isRootScroller(scroller)) return { top: 0, bottom: window.innerHeight, zoom: 1 };
  const zoom = effectiveZoom(scroller);
  const top = scroller.getBoundingClientRect().top;
  return { top, bottom: top + scroller.clientHeight * zoom, zoom };
}

/**
 * The band of `column` the viewport shows, in the column's layout pixels, the
 * space its page offsets are in. `scroller` is null for the root scroller.
 */
export function viewportColumnBand(
  scroller: HTMLElement | null,
  column: HTMLElement
): { top: number; bottom: number } {
  const viewport = scroller ? scrollViewport(scroller) : { top: 0, bottom: window.innerHeight };
  const origin = column.getBoundingClientRect().top;
  const zoom = effectiveZoom(column);
  return { top: (viewport.top - origin) / zoom, bottom: (viewport.bottom - origin) / zoom };
}

/** The `scrollTop` change that brings `[top, bottom]`, in client pixels, `margin` layout pixels inside the viewport. */
export function scrollIntoViewDelta(
  viewport: ScrollViewport,
  top: number,
  bottom: number,
  margin: number
): number {
  const start = (top - viewport.top) / viewport.zoom;
  const end = (bottom - viewport.top) / viewport.zoom;
  const height = (viewport.bottom - viewport.top) / viewport.zoom;
  if (start < margin) return start - margin;
  if (end > height - margin) return end - height + margin;
  return 0;
}

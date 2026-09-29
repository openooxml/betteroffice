import { effectiveZoom } from '@betteroffice/docx/layout/render';

/** What a scroller shows, measured the way each caller measured it before zoom was handled. */
export interface ScrollViewport {
  /** Client top of the scroller's box; the window's for the root. */
  top: number;
  /** Client bottom of the scroller's box; the window's for the root. */
  bottom: number;
  /** Layout height the scroller shows its content in: `clientHeight`, or the window's. */
  height: number;
  /** Client pixels per unit of `scrollTop`: the root scrolls in viewport units. */
  zoom: number;
}

function isRootScroller(scroller: Element): boolean {
  return (
    typeof document !== 'undefined' &&
    scroller === (document.scrollingElement ?? document.documentElement)
  );
}

export function scrollViewport(scroller: HTMLElement): ScrollViewport {
  if (isRootScroller(scroller)) {
    return { top: 0, bottom: window.innerHeight, height: window.innerHeight, zoom: 1 };
  }
  const rect = scroller.getBoundingClientRect();
  return {
    top: rect.top,
    bottom: rect.bottom,
    height: scroller.clientHeight,
    zoom: effectiveZoom(scroller),
  };
}

/**
 * The viewport band over `column`, in the column's layout pixels measured from
 * the client origin: the band starts `top - columnTop` down the column.
 * `scroller` is null for the root scroller.
 */
export function viewportColumnBand(
  scroller: HTMLElement | null,
  column: HTMLElement
): { columnTop: number; top: number; height: number } {
  const viewport = scroller
    ? scrollViewport(scroller)
    : { top: 0, height: window.innerHeight, zoom: 1 };
  const zoom = effectiveZoom(column);
  return {
    columnTop: column.getBoundingClientRect().top / zoom,
    top: viewport.top / zoom,
    height: (viewport.height * viewport.zoom) / zoom,
  };
}

/**
 * The `scrollTop` change that brings the client-pixel span `[top, bottom]`
 * `margin` layout pixels inside the scroller's box.
 */
export function scrollIntoViewDelta(
  viewport: ScrollViewport,
  top: number,
  bottom: number,
  margin: number
): number {
  const { zoom } = viewport;
  if (top < viewport.top + margin * zoom) return (top - viewport.top) / zoom - margin;
  if (bottom > viewport.bottom - margin * zoom) return (bottom - viewport.bottom) / zoom + margin;
  return 0;
}

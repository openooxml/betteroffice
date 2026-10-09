import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import { viewportColumnBand } from '../internals/viewportBand';
import {
  remotePresencePageWindow,
  type RemotePresencePageMetrics,
  type RemotePresencePageWindow,
} from './remotePresenceGeometry';

/**
 * Reports the pages in view, and one either side, a frame after anything that
 * can change them: scrolling any scroller, resizing, pages coming into view or
 * being replaced. `moved` is true when pages may have moved or been replaced
 * under an unchanged window. Returns the disposer.
 */
export function observePageWindow(
  host: HTMLElement,
  metrics: RemotePresencePageMetrics,
  onWindow: (pageWindow: RemotePresencePageWindow | null, moved: boolean) => void,
  maxPages = Infinity
): () => void {
  const scrollsWindow = (element: HTMLElement) =>
    element === document.scrollingElement || element === document.documentElement;
  let frame: number | null = null;
  let moved = true;
  const schedule = () => {
    if (frame !== null) return;
    frame = requestAnimationFrame(() => {
      frame = null;
      update();
    });
  };
  const reflow = () => {
    moved = true;
    schedule();
  };
  // The scroller is looked up again on every update: an editor shown after
  // mounting, or an ancestor that starts scrolling, changes it without any
  // event on the old one.
  let scrollParent: HTMLElement | null = null;
  let scrollerResized: ResizeObserver | null = null;
  const update = () => {
    const next = findVerticalScrollParentOrRoot(host);
    if (next !== scrollParent) {
      scrollerResized?.disconnect();
      scrollParent = next;
      scrollerResized = scrollsWindow(next) ? null : new ResizeObserver(reflow);
      scrollerResized?.observe(next);
    }
    const usesWindow = scrollsWindow(next);
    const column = host.firstElementChild as HTMLElement | null;
    const band = column ? viewportColumnBand(usesWindow ? null : next, column) : null;
    const pageWindow = band
      ? remotePresencePageWindow(
          metrics,
          band.columnTop,
          band.top,
          band.top + band.height,
          maxPages
        )
      : null;
    const wasMoved = moved;
    moved = false;
    onWindow(pageWindow, wasMoved);
  };
  // Scroll events do not bubble; capturing them on the document sees every
  // scroller, the window's included.
  document.addEventListener('scroll', schedule, { capture: true, passive: true });
  const hostResized = new ResizeObserver(reflow);
  hostResized.observe(host);
  // Layout elsewhere can move pages into view without scrolling or resizing.
  const pagesShown =
    typeof IntersectionObserver === 'function' ? new IntersectionObserver(reflow) : null;
  const observePages = () => {
    pagesShown?.disconnect();
    host.querySelectorAll('.canvas-page').forEach((page) => pagesShown?.observe(page));
  };
  observePages();
  // The page surfaces are replaced without a display-list change when the
  // canvas falls back from worker rendering.
  const pagesReplaced =
    typeof MutationObserver === 'function'
      ? new MutationObserver(() => {
          const column = host.firstElementChild;
          if (column) pagesReplaced?.observe(column, { childList: true });
          observePages();
          reflow();
        })
      : null;
  pagesReplaced?.observe(host, { childList: true });
  if (host.firstElementChild) pagesReplaced?.observe(host.firstElementChild, { childList: true });
  window.addEventListener('resize', reflow);
  host.addEventListener('transitionend', reflow);
  update();
  return () => {
    document.removeEventListener('scroll', schedule, { capture: true });
    scrollerResized?.disconnect();
    hostResized.disconnect();
    pagesShown?.disconnect();
    pagesReplaced?.disconnect();
    window.removeEventListener('resize', reflow);
    host.removeEventListener('transitionend', reflow);
    if (frame !== null) cancelAnimationFrame(frame);
  };
}

import type { Layout } from '@betteroffice/docx/layout/pagination';
import { VIEWPORT_PADDING_BOTTOM, VIEWPORT_PADDING_TOP } from './styles';

export function runAfterFrames(fn: () => void, signal: AbortSignal): void {
  if (signal.aborted) return;
  const id1 = requestAnimationFrame(() => {
    if (signal.aborted) return;
    const id2 = requestAnimationFrame(() => {
      if (signal.aborted) return;
      const id3 = requestAnimationFrame(() => {
        if (!signal.aborted) fn();
      });
      signal.addEventListener('abort', () => cancelAnimationFrame(id3), { once: true });
    });
    signal.addEventListener('abort', () => cancelAnimationFrame(id2), { once: true });
  });
  signal.addEventListener('abort', () => cancelAnimationFrame(id1), { once: true });
}

/** Min-height of the zoom/viewport wrapper: top + bottom padding plus the page stack. */
export function viewportMinHeightPx(layout: Layout, pageGap: number): number {
  const n = layout.pages.length;
  const pagesHeight = layout.pages.reduce((sum, page) => sum + page.size.h, 0);
  return (
    pagesHeight + Math.max(0, n - 1) * pageGap + VIEWPORT_PADDING_TOP + VIEWPORT_PADDING_BOTTOM
  );
}

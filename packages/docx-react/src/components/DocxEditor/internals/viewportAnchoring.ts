export type LayoutUpdateOrigin = 'local' | 'remote';

export interface ViewportAnchorSnapshot {
  viewportOffset: number;
  scrollTopSnapshot: number;
}

export function computeViewportAnchoredScrollTop(
  anchor: ViewportAnchorSnapshot,
  nextTargetTop: number | null,
  maxScrollTop: number
): number {
  const requested =
    nextTargetTop == null ? anchor.scrollTopSnapshot : nextTargetTop - anchor.viewportOffset;
  return Math.min(Math.max(0, requested), Math.max(0, maxScrollTop));
}

export function shouldScrollCaretIntoView(
  layoutUpdateOrigin: LayoutUpdateOrigin,
  selectionChanged: boolean
): boolean {
  return layoutUpdateOrigin === 'local' || selectionChanged;
}

import type { RetainedFrame } from '../layout/render/frameDelta';
import type { YrsResidentCaretRect, YrsResidentCaretSnapshot } from './index';

/** Resolved styling for the worker-painted caret (the worker cannot read CSS vars). */
export interface ResidentCaretPaintStyle {
  color: string;
  /** CSS-px stroke width, matching the DOM caret's un-zoomed width. */
  width: number;
}

/** Device-px caret rect: geometry scales by dpr*zoom, stroke width by dpr only
 * (the DOM caret keeps its CSS-px width across zoom levels). */
export function residentCaretDeviceRect(
  rect: Pick<YrsResidentCaretRect, 'x' | 'y' | 'height'>,
  style: ResidentCaretPaintStyle,
  devicePixelRatio: number,
  zoom: number
): { x: number; y: number; width: number; height: number } {
  const scale = devicePixelRatio * zoom;
  return {
    x: rect.x * scale,
    y: rect.y * scale,
    width: Math.max(1, style.width * devicePixelRatio),
    height: rect.height * scale,
  };
}

export function residentCaretSnapshotForFrame(
  snapshot: YrsResidentCaretSnapshot,
  frame: RetainedFrame
): YrsResidentCaretSnapshot | null {
  if (!Number.isSafeInteger(snapshot.frameEpoch) || snapshot.frameEpoch !== frame.frameEpoch) {
    return null;
  }
  const rect = snapshot.caretRect;
  if (!rect) return snapshot;
  if (
    !Number.isInteger(rect.pageIndex) ||
    rect.pageIndex < 0 ||
    !Number.isFinite(rect.x) ||
    !Number.isFinite(rect.y) ||
    !Number.isFinite(rect.height) ||
    rect.height < 0 ||
    !/^[1-9]\d*$/u.test(rect.pageId)
  ) {
    return null;
  }
  const page = frame.pages[rect.pageIndex];
  return page?.pageId.toString() === rect.pageId ? snapshot : null;
}

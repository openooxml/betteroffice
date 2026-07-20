import type { RetainedFrame } from '../layout/render/frameDelta';
import type { YrsResidentCaretSnapshot } from './index';

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

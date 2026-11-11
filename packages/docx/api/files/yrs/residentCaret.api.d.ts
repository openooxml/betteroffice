// @betteroffice/docx declaration yrs/residentCaret.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
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
export declare function residentCaretDeviceRect(rect: Pick<YrsResidentCaretRect, 'x' | 'y' | 'height'>, style: ResidentCaretPaintStyle, devicePixelRatio: number, zoom: number): {
    x: number;
    y: number;
    width: number;
    height: number;
};
export declare function residentCaretSnapshotForFrame(snapshot: YrsResidentCaretSnapshot, frame: RetainedFrame): YrsResidentCaretSnapshot | null;

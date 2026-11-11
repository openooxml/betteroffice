// @betteroffice/docx declaration yrs/residentControls.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxContentControl } from './contentControls';
import type { YrsContentControlValue } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { DisplayListQueries } from '../layout/render/displayListQueries';
export interface ResidentControlPosition {
    story: string;
    index: number;
    encoded: number[];
}
export interface ResidentControlTarget {
    control: DocxContentControl;
    at: ResidentControlPosition;
}
export interface ResidentControlWidget {
    kind: 'checkbox' | 'dropdown' | 'date' | 'repeatingItem';
    checked?: boolean;
    value?: string;
    items?: {
        displayText: string;
        value: string;
    }[];
}
export interface ResidentControlSnapshot {
    version: string;
    frameEpoch: number | null;
    layoutRevision: number;
    controls: {
        handle: string;
        control: DocxContentControl;
        location: {
            story: string;
            index: number;
            rootStory: string;
            position: number | null;
        };
        widget: ResidentControlWidget | null;
        rects: {
            pageId: string;
            pageIndex: number;
            x: number;
            y: number;
            width: number;
            height: number;
        }[];
    }[];
}
/**
 * The controls of the current version with their widgets and, given the frame that shows it, their page rects:
 * the layout's `sdt@` group, else the painted content of an inline control the layout groups no widget for.
 */
export declare function readControlHandles(engine: ResidentEngineSession, frame?: RetainedFrame, layoutRevision?: number, queries?: () => DisplayListQueries): ResidentControlSnapshot;
export declare function resolveControlHandle(engine: ResidentEngineSession, handle: string): ResidentControlTarget | null;
export declare function validControlValue(value: unknown): value is YrsContentControlValue;

// @betteroffice/docx declaration yrs/residentObjectGeometry.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsCellLoc, YrsLoc, YrsStickyPosition, YrsTableRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandSelection } from './residentCommandProtocol';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { DisplayListRect } from '../layout/render/displayListQueries';
import type { DisplayPrimitive } from '../layout/render/displayList';
export interface ResidentObjectRect {
    pageId: string;
    pageIndex: number;
    x: number;
    y: number;
    width: number;
    height: number;
    /** Table cells only: this box is a header row's copy on a continuation page. */
    repeatedHeader?: boolean;
}
export interface ResidentObjectGeometry {
    handle: string;
    kind: 'text' | 'cell' | 'image';
    story: string;
    rects: ResidentObjectRect[];
    /** Images only: anchored to the page rather than flowing with the text. */
    floating?: boolean;
    /** Floating images only: painted behind the text, so text over them takes the pointer. */
    behindText?: boolean;
    /** Images only: the picture's aspect ratio is locked. */
    noChangeAspect?: boolean;
    rotation?: number;
    /** Images only: the worker selection covers exactly this picture's embed. */
    selected?: boolean;
    /** Image paint order within the page. */
    paintOrder?: number;
    /** Image anchored inside a cell or text box. */
    nested?: boolean;
    /** Cells only: where the cell sits in its table's grid. */
    cell?: ResidentCellPlacement;
}
export interface ResidentCellPlacement {
    /** Identifies the table within the document: its story and index there. */
    table: string;
    row: number;
    column: number;
    rowspan: number;
    colspan: number;
    /** Grid columns of the whole table. */
    columns: number;
    /** Whether the cell's boxes are its real edges rather than an estimate from its content. */
    exact: boolean;
}
export declare function residentCellPlacement(table: {
    story: string;
    tableIndex: number;
    cells: ReadonlyArray<{
        column: number;
        colspan: number;
    }>;
}, cell: {
    row: number;
    column: number;
    rowspan: number;
    colspan: number;
}, exact: boolean): ResidentCellPlacement;
export declare function isFloatingImagePayload(payload: Record<string, unknown>): boolean;
/** Whether layout positions the picture by its anchor; other wraps (Top and Bottom, for one) flow with the text and ignore the offsets. */
export declare function isAnchoredImagePayload(payload: Record<string, unknown>): boolean;
/** Whether `selection` covers exactly the one embed at `at`. */
export declare function embedSelected(selection: DocxCommandSelection | null, at: YrsLoc): boolean;
export declare function imagePlacement(payload: Record<string, unknown>): Pick<ResidentObjectGeometry, 'floating' | 'behindText' | 'noChangeAspect' | 'rotation'>;
export interface ResidentFrameGeometry {
    version: string;
    frameEpoch: number;
    layoutRevision: number;
    selection: ResidentObjectRect[];
    objects: ResidentObjectGeometry[];
    unbuiltPages: number[];
}
interface ObjectTarget {
    kind: ResidentObjectGeometry['kind'];
    position: YrsStickyPosition;
    paraId: string;
    story: string;
}
export declare function resetResidentObjectHandles(engine: ResidentEngineSession): void;
export declare function resolveResidentObject(engine: ResidentEngineSession, handle: string): (YrsLoc & {
    kind: ObjectTarget['kind'];
    cell?: YrsCellLoc;
}) | null;
/** Page-local boxes of a multi-cell range; null when no multi-cell range is active. */
export declare function readResidentCellSelectionRects(engine: ResidentEngineSession, frame: RetainedFrame, range: YrsTableRange | null): DisplayListRect[] | null;
export declare function residentImagePaintOrders(frame: RetainedFrame): Map<DisplayPrimitive, number>;
export declare function readResidentFrameGeometry(engine: ResidentEngineSession, frame: RetainedFrame, selection: DocxCommandSelection | null, layoutRevision: number): Promise<ResidentFrameGeometry>;
export {};

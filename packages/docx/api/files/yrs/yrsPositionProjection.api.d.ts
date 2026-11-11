// @betteroffice/docx declaration yrs/yrsPositionProjection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsCellLoc, YrsLoc, YrsSession, YrsStorySegment } from './index';
import type { PointPosition } from '../plugin-api';
import { type YrsInputPositionMap } from './inputPositionMap';
type PositionReader = Pick<YrsSession, 'hasStory' | 'storySegments'>;
export type YrsLocProjection = Pick<YrsPositionProjection, 'positionForLoc'>;
export type YrsPositionOutline = Record<string, {
    contentStart: number;
    size: number;
    /** `length` counts display units; `controls` lists inline controls as `[story offset, display units]`. */
    paragraphs: Array<{
        paraId: string;
        displayStart: number;
        length: number;
        leading: number;
        controls?: InlineControls;
    }>;
}>;
/** Inline controls of a paragraph as `[story offset within the content, display units]`, in order. */
type InlineControls = ReadonlyArray<readonly [number, number]>;
/** Display units of an inline control with `payload`: its content and its two boundaries, as the layout counts them. */
export declare function inlineControlDisplaySize(payload: Record<string, unknown>): number;
export declare function createYrsLocProjectionFromOutline(outline: YrsPositionOutline): YrsLocProjection;
/** @internal */
export interface YrsProjectedNode {
    kind: string;
    start: number;
    nodeSize: number;
    attrs: Record<string, unknown>;
    story: string;
}
/** @internal */
export interface YrsProjectedTable extends YrsProjectedNode {
    kind: 'table';
    tableIndex: number;
    widthsTwips: number[];
    rowCount: number;
    cells: YrsProjectedCell[];
}
/** @internal */
export interface YrsProjectedCell {
    row: number;
    column: number;
    rowspan: number;
    colspan: number;
    story: string;
    start: number;
    contentStart: number;
    nodeSize: number;
    attrs: Record<string, unknown>;
}
/** @internal */
export interface YrsPointerProjectionTarget {
    story: string;
    displayPosition: number;
    cell?: YrsCellLoc;
}
/** @internal */
export interface YrsStorySegmentSource {
    segments(story: string): YrsStorySegment[];
}
/** @internal */
export declare function createYrsPositionProjection(session: PositionReader, rootStory: string, source?: YrsStorySegmentSource): YrsPositionProjection | null;
/** @internal */
export declare function projectYrsDisplayPosition(position: number | PointPosition, getProjection: (rootStory: string) => YrsPositionProjection | null, session?: object | null): YrsPointerProjectionTarget | null;
/** @internal */
export declare class YrsPositionProjection {
    readonly rootStory: string;
    private readonly stories;
    private readonly nodes;
    private readonly tables;
    private readonly blockBoundaries;
    private readonly source;
    private readonly paragraphsById;
    constructor(session: Pick<YrsSession, 'storySegments'>, rootStory: string, source?: YrsStorySegmentSource);
    get size(): number;
    nodeAt(position: number): YrsProjectedNode | null;
    allTables(): readonly YrsProjectedTable[];
    tableAtPosition(position: number): YrsProjectedTable | null;
    tableAtStart(start: number): YrsProjectedTable | null;
    tableStartForLoc(at: YrsCellLoc): number | null;
    cellPosition(tableStart: number, row: number, column: number): number | null;
    targetAt(position: number): YrsPointerProjectionTarget;
    positionForLoc(loc: YrsLoc): number | null;
    /** The location of a block embed's start or end at `position`, in the offsets of the paragraph after it. */
    blockBoundaryAt(position: number): YrsLoc | null;
    private setBlockBoundary;
    /** First paragraph for each paraId. */
    private paragraphIndex;
    bookmarkPosition(name: string): number | null;
    private buildStory;
    private inputPositionAt;
    private buildTable;
}
/** @internal */
export declare function yrsLocToProjectedDisplayPosition(reader: Pick<YrsSession, 'hasStory' | 'paragraphSpans'>, projectionFor: (rootStory: string) => YrsLocProjection | null, loc: YrsLoc, activeRootStory?: string, inputMap?: (story: string) => YrsInputPositionMap | null): number | null;
export {};

// @betteroffice/docx declaration layout/render/yrsSidebarProjection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsSession, YrsStorySegment } from '../../yrs';
type SidebarReader = Pick<YrsSession, 'storyIds' | 'storySegments' | 'version' | 'paragraphs' | 'locateParagraph'>;
/** A yrs location projected into the position space used by the display list. */
export interface YrsSidebarDisplayPoint {
    story: string;
    position: number;
    /** Header/footer relationship id; absent for body and nested body stories. */
    hfRid?: string;
}
/** Read-only converter used by the yrs-backed sidebar data sources. */
export interface YrsSidebarProjection {
    locToDisplayPoint(loc: YrsLoc): YrsSidebarDisplayPoint | null;
    storyOffsetToDisplayPoint(story: string, offset: number): YrsSidebarDisplayPoint | null;
}
/**
 * Where a sidebar projection reads story segments. `segments` must answer for the session's
 * current state on every call, as `YrsSession.storySegments` does; the arrays it returns are shared
 * and never mutated. A projection remembers the last source given for a session and reads
 * through it for the session's lifetime.
 */
export interface YrsStorySegmentSource {
    segments(story: string): YrsStorySegment[];
}
/**
 * Convert a string yrs identity into the numeric id used by the existing
 * layout/sidebar contract. This is the same UTF-8 FNV-1a projection used by
 * the yrs render and Document bridges; numeric OOXML ids pass through.
 */
export declare function yrsIdToNumericId(value: string): number;
/** @internal */
export declare function hasCachedYrsSidebarProjection(session: SidebarReader): boolean;
/**
 * Build a lazy projection from live yrs stories to display positions.
 * The canonical yrs segment stream supplies paragraph/atom units; table-cell
 * and block-SDT stories are recursively sized so container tokens are included.
 * A session gets the same projection back until its document changes.
 */
export declare function createYrsSidebarProjection(session: SidebarReader, source?: YrsStorySegmentSource): YrsSidebarProjection;
export {};

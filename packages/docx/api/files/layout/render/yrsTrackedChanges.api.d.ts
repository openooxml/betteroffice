// @betteroffice/docx declaration layout/render/yrsTrackedChanges.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsRevisionInfo } from '../../yrs';
import type { TrackedChangeEntry } from '../../utils/comments';
import { type YrsSidebarProjection } from './yrsSidebarProjection';
/** Tracked-change sidebar data derived from the authoritative Yrs session. */
export interface TrackedChangesResult {
    entries: TrackedChangeEntry[];
    commentToRevision: Map<number, number>;
}
/** Project Yrs revisions onto the display-position sidebar contract. */
export declare function extractTrackedChangesFromYrs(revisions: readonly YrsRevisionInfo[], projection: YrsSidebarProjection): TrackedChangesResult;

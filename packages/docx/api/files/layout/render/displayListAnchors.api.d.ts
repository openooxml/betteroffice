// @betteroffice/docx declaration layout/render/displayListAnchors.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Comment / tracked-change anchor derivation over the display list, shared
 * by both adapters' sidebar plumbing on the canvas-renderer path.
 *
 * `visitAnchorKeys` walks the document and reports every comment /
 * tracked-change anchor as `(key, pos)` — shared between each adapter's
 * layout-math derivation (DOM painter path) and the display-list derivation
 * here, so both register identical keys in identical order.
 *
 */
import type { YrsRevisionInfo, YrsSession } from '../../yrs';
import type { DisplayListQueries, DisplayListRect } from './displayListQueries';
import { type YrsSidebarDisplayPoint, type YrsSidebarProjection } from './yrsSidebarProjection';
/** Structural view shape used only by the legacy display-list anchor walker. */
export interface AnchorEditorView {
    readonly state: {
        readonly doc: {
            descendants(callback: (node: any, pos: number) => boolean | void): void;
        };
        readonly schema: {
            readonly marks: Record<string, any>;
        };
    };
}
/**
 * Walk the document and report every comment / tracked-change anchor as
 * `(key, pos)`. `register` is called once per occurrence; callers dedupe.
 */
export declare function visitAnchorKeys(pmView: AnchorEditorView, register: (key: string, pos: number) => void): void;
/**
 * Display-list anchor derivation: anchor Ys come from Rust `range_rects`
 * queries over the display list (page-local px) plus the canvas page stack
 * math, instead of layout-fragment scans in the DOM viewport's coordinate
 * space. Returns `.canvas-pages`-host Y offsets keyed by
 * `comment-{id}` / `revision-{revisionId}`.
 */
export declare function computeAnchorPositionsFromDisplayList(pmView: AnchorEditorView | null, queries: DisplayListQueries, projectY?: (rect: DisplayListRect) => number | null): Map<string, number>;
export interface YrsHeaderFooterRegions {
    get(rId: string): 'header' | 'footer' | undefined;
}
/**
 * yrs-authoritative sidebar anchors. Comment coverage comes from sticky yrs
 * side-map positions and tracked-change anchors come from `listRevisions()`;
 * neither path reads an `AnchorEditorView` or its document.
 */
export declare function computeAnchorPositionsFromYrs(session: YrsSession, commentIds: Iterable<string | number>, revisions: readonly YrsRevisionInfo[], projection: YrsSidebarProjection, queries: DisplayListQueries, hfRegions?: YrsHeaderFooterRegions, projectY?: (rect: DisplayListRect) => number | null): Map<string, number>;
export declare function anchorPositionsFromPoints(points: Iterable<readonly [key: string, point: YrsSidebarDisplayPoint | null]>, queries: DisplayListQueries, hfRegions?: YrsHeaderFooterRegions, projectY?: (rect: DisplayListRect) => number | null): Map<string, number>;
export interface DisplayListHfAnchorSource {
    region: 'header' | 'footer';
    rId: string;
    view: AnchorEditorView;
}
/** Merge header/footer comment and revision anchors into an existing map. */
export declare function mergeHfAnchorPositionsFromDisplayList(sources: Iterable<DisplayListHfAnchorSource>, queries: DisplayListQueries, positions: Map<string, number>, projectY?: (rect: DisplayListRect) => number | null): Map<string, number>;

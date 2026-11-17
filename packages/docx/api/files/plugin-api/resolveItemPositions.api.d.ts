// @betteroffice/docx declaration plugin-api/resolveItemPositions.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Pure layout algorithm: given a list of sidebar items + their
 * anchor sources, return a list of resolved Y positions with
 * collision avoidance applied. Lifted from React's
 * components/sidebar/resolveItemPositions.ts so the Vue adapter
 * can call it directly.
 *
 * The algorithm walks the items in three priority stages:
 *   1. explicit fixedY (already in scroll-container coords)
 *   2. anchorKey lookup in anchorPositions (layout/pagination coords)
 *   3. DOM-based anchor lookup via renderedDomContext.getRectsForRange
 *   4. last-known cache so cards don't pop out during transient layout
 *
 * Then it sorts by target Y and pushes overlapping cards down by
 * their height + MIN_CARD_GAP.
 * @packageDocumentation
 * @public
 */
import type { RenderedDomContext, SidebarItem } from './types';
/** Generic sidebar item with optional fixedY / estimatedHeight. */
export interface ResolvableSidebarItem extends SidebarItem {
    fixedY?: number;
    estimatedHeight?: number;
}
export interface ResolvedPosition<T extends ResolvableSidebarItem = ResolvableSidebarItem> {
    item: T;
    y: number;
}
export declare function resolveItemPositions<T extends ResolvableSidebarItem>(items: T[], anchorPositions: Map<string, number>, renderedDomContext: RenderedDomContext | null, zoom: number, cardHeights: Map<string, number>, lastKnown: Map<string, number>): ResolvedPosition<T>[];

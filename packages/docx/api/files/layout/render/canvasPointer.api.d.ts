// @betteroffice/docx declaration layout/render/canvasPointer.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Client-coordinate → display-list resolution for the canvas renderer's
 * pointer path. Framework-free: both the React and the Vue adapter route
 * their pointer events through this resolver.
 *
 * The pointer hooks map the client point onto one of the
 * `<canvas data-page-index>` pages, convert it to page-local px (the
 * display list's unit space), and ask the Rust `hit_test_regions` query for
 * the region + doc position.
 *
 */
import type { DisplayListQueries, DisplayListRegionHit } from './displayListQueries';
export interface CanvasPointHit {
    pageIndex: number;
    /** page-local px, the display list's unit space */
    x: number;
    y: number;
    /** region-aware hit from the Rust query (null when the engine isn't ready) */
    hit: DisplayListRegionHit | null;
}
export interface DisplayPageHostOptions {
    /** Vertical space between page shells when the host has no live canvases. */
    pageGap?: number;
    /** Space before the first page when the host has no live canvases. */
    paddingTop?: number;
}
export interface DisplayPageClientRect {
    left: number;
    top: number;
    right: number;
    bottom: number;
    width: number;
    height: number;
}
/**
 * The CSS `zoom` `element` renders at, its ancestors' included. Client rects
 * carry it; layout sizes and an element's scroll offsets do not. Where
 * `currentCSSZoom` is unavailable, zoom set inside a closed shadow tree the
 * element is slotted through cannot be read and is left out.
 */
export declare function effectiveZoom(element: Element): number;
/**
 * The page canvases a renderer mounts under its pages host, so page lookups
 * read them instead of searching a subtree that also holds every page's
 * accessibility mirror. A renderer adds each `<canvas data-page-index>` it
 * mounts, deletes it on unmount, invalidates the registry whenever it may have
 * moved or renumbered them, and binds the registry to the host with
 * {@link bindDisplayPageRegistry}.
 */
export declare class DisplayPageRegistry {
    private readonly canvases;
    private ordered;
    private byIndex;
    private materializer;
    /**
     * Lets a renderer that keeps page DOM only near the viewport build the
     * DOM of other pages when something needs it; see {@link materializeDisplayPages}.
     */
    setMaterializer(materializer: ((pageIndices: readonly number[]) => void) | null): void;
    /** Builds the DOM of `pageIndices` now, where the renderer keeps it only for some pages. */
    materialize(pageIndices: readonly number[]): void;
    add(canvas: HTMLCanvasElement): void;
    delete(canvas: HTMLCanvasElement): void;
    /** Forgets the order and indices read so far, after the renderer moved or renumbered canvases. */
    invalidate(): void;
    /** The registered canvases in document order. */
    all(): HTMLCanvasElement[];
    /** The first canvas in document order that shows `pageIndex`. */
    canvas(pageIndex: number): HTMLCanvasElement | null;
    private indexPages;
}
/** Binds `registry` to the pages `host` it renders into, or unbinds it with null. */
export declare function bindDisplayPageRegistry(host: HTMLElement, registry: DisplayPageRegistry | null): void;
/** Builds the accessibility DOM of `pageIndices` under `host` before a DOM query reads it. */
export declare function materializeDisplayPages(host: HTMLElement, pageIndices: readonly number[]): void;
/** The first `<canvas data-page-index>` under `host` for `pageIndex`. */
export declare function displayPageCanvas(host: HTMLElement, pageIndex: number): HTMLCanvasElement | null;
/** The canvas page elements inside a pages host, in document order. */
export declare function displayPageCanvases(host: HTMLElement): HTMLCanvasElement[];
/** Resolves a display-list page rectangle from canvas or DOM page hosts. */
export declare function resolveDisplayPageClientRect(host: HTMLElement, queries: DisplayListQueries, pageIndex: number, options?: DisplayPageHostOptions): DisplayPageClientRect | null;
/**
 * Resolve a client point against the canvas pages.
 *
 * Containment picks the page; with `clampToNearestPage` (drag-selection) a
 * point outside every page snaps to the vertically nearest one and clamps
 * into its bounds, so drags keep extending past page edges — the analogue of
 * the DOM path's nearest-span snapping. Coordinates convert through the
 * page's CSS rect vs its display-list size, so any ancestor scale transform
 * is factored out.
 */
export declare function resolveCanvasPoint(host: HTMLElement, queries: DisplayListQueries, clientX: number, clientY: number, options?: {
    clampToNearestPage?: boolean;
    pageGap?: number;
    paddingTop?: number;
}): CanvasPointHit | null;

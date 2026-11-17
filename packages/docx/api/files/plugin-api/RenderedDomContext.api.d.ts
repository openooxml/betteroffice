// @betteroffice/docx declaration plugin-api/RenderedDomContext.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * RenderedDomContext Implementation
 *
 * Provides position mapping over the canvas renderer's output. Geometry is
 * served by the Rust display-list queries when available; DOM fallbacks read
 * the accessibility mirror's data-doc-start/data-doc-end span contract.
 * @packageDocumentation
 * @public
 */
import type { RenderedDomContext, PositionCoordinates, PointPosition } from './types';
import type { DisplayListQueries, DisplayListRect } from '../layout/render/displayListQueries';
/**
 * Options controlling how a {@link RenderedDomContextImpl} resolves geometry.
 *
 * @public
 */
export interface RenderedDomContextOptions {
    /** Immutable display-list query source. Omit for the mirror-DOM backend. */
    displayListQueries?: DisplayListQueries;
    /** Page-local display-list geometry → pages-container coordinates. */
    projector?: DisplayListPageProjector;
}
export interface ProjectedDisplayListRect {
    x: number;
    y: number;
    width: number;
    height: number;
}
/** Canvas-host projection contract kept separate from document geometry. */
export interface DisplayListPageProjector {
    projectRect(rect: DisplayListRect): ProjectedDisplayListRect | null;
    getPageBounds(pageIndex: number): ProjectedDisplayListRect | null;
}
/**
 * Project through the live canvas for each page. Reading its client rect is a
 * host-coordinate transform only; document geometry stays display-list-owned.
 */
export declare function createCanvasHostProjector(pagesContainer: HTMLElement, queries: DisplayListQueries, zoom?: number): DisplayListPageProjector;
/**
 * Implementation of RenderedDomContext.
 *
 * This class provides position mapping between display-list document
 * positions and pixel coordinates in the rendered output. Geometry comes
 * from the Rust display-list queries; the DOM fallbacks use the
 * data-doc-start and data-doc-end attributes the accessibility mirror
 * emits (see {@link RenderedDomContextOptions}).
 */
export declare class RenderedDomContextImpl implements RenderedDomContext {
    pagesContainer: HTMLElement;
    zoom: number;
    private readonly queries;
    private readonly projector;
    constructor(pagesContainer: HTMLElement, zoom?: number, options?: RenderedDomContextOptions);
    /** Client pixels per unit of this context: its zoom times the container's ancestor CSS zoom. */
    private clientScale;
    /**
     * Run spans (`data-doc-start`/`data-doc-end`) in reading order, deep-queried
     * from the a11y mirror subtree nested under `.canvas-pages`.
     */
    private spanEntries;
    /** Empty-paragraph runs, deep-queried from the a11y mirror. */
    private emptyRuns;
    /**
     * Get pixel coordinates for a display position.
     * Uses the browser's text rendering via Range API for precise positioning.
     */
    getCoordinatesForPosition(position: number): PositionCoordinates | null;
    getPositionAtPoint(clientX: number, clientY: number): PointPosition | null;
    /**
     * Find DOM elements that overlap with a display-position range.
     */
    findElementsForRange(from: number, to: number): Element[];
    /**
     * Get bounding rectangles for a range of text.
     * Handles line wraps by returning multiple rects.
     */
    getRectsForRange(from: number, to: number): Array<{
        x: number;
        y: number;
        width: number;
        height: number;
    }>;
    getCoordinatesForHfPosition(region: 'header' | 'footer', rId: string, position: number, pageIndex: number): PositionCoordinates | null;
    getRectsForHfRange(region: 'header' | 'footer', rId: string, from: number, to: number, pageIndex: number): ProjectedDisplayListRect[];
    getPageBounds(pageIndex: number): ProjectedDisplayListRect | null;
    /**
     * Get the offset of the pages container from its parent viewport.
     * This is needed for positioning overlays that are rendered in the
     * viewport container rather than directly in the pages container.
     */
    getContainerOffset(): {
        x: number;
        y: number;
    };
}
/**
 * Create a RenderedDomContext for a pages container element.
 *
 * @param pagesContainer - The `.canvas-pages` host holding the rendered pages
 * @param zoom - Current zoom level (default 1)
 * @param options - query/projector backend options. Omit for the mirror-DOM
 *   fallback.
 */
export declare function createRenderedDomContext(pagesContainer: HTMLElement, zoom?: number, options?: RenderedDomContextOptions): RenderedDomContext & Required<Pick<RenderedDomContext, 'getPositionAtPoint'>>;

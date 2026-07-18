/**
 * RenderedDomContext Implementation
 *
 * Provides position mapping over the canvas renderer's output. Geometry is
 * served by the Rust display-list queries when available; DOM fallbacks read
 * the accessibility mirror's data-doc-start/data-doc-end span contract.
 * @packageDocumentation
 * @public
 */

import type { RenderedDomContext, PositionCoordinates } from './types';
import type { DisplayListQueries, DisplayListRect } from '../layout/render/displayListQueries';
import { resolveDisplayPageClientRect } from '../layout/render/canvasPointer';

/** One data-doc-* bearing run span in the a11y mirror, with parsed positions. */
interface MirrorSpanEntry {
  el: HTMLElement;
  start: number;
  end: number;
}

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
export function createCanvasHostProjector(
  pagesContainer: HTMLElement,
  queries: DisplayListQueries,
  zoom = 1
): DisplayListPageProjector {
  const safeZoom = zoom > 0 ? zoom : 1;
  const projectRect = (rect: DisplayListRect): ProjectedDisplayListRect | null => {
    const pageSize = queries.pageSize(rect.pageIndex);
    const canvasRect = resolveDisplayPageClientRect(pagesContainer, queries, rect.pageIndex);
    if (!canvasRect || !pageSize || pageSize.width <= 0 || pageSize.height <= 0) return null;
    const containerRect = pagesContainer.getBoundingClientRect();
    const scaleX = canvasRect.width / pageSize.width;
    const scaleY = canvasRect.height / pageSize.height;
    return {
      x: (canvasRect.left - containerRect.left + rect.x * scaleX) / safeZoom,
      y: (canvasRect.top - containerRect.top + rect.y * scaleY) / safeZoom,
      width: (rect.width * scaleX) / safeZoom,
      height: (rect.height * scaleY) / safeZoom,
    };
  };
  return {
    projectRect,
    getPageBounds: (pageIndex: number) => {
      const bounds = queries.pageBounds(pageIndex);
      return bounds ? projectRect(bounds) : null;
    },
  };
}

/**
 * Implementation of RenderedDomContext.
 *
 * This class provides position mapping between display-list document
 * positions and pixel coordinates in the rendered output. Geometry comes
 * from the Rust display-list queries; the DOM fallbacks use the
 * data-doc-start and data-doc-end attributes the accessibility mirror
 * emits (see {@link RenderedDomContextOptions}).
 */
export class RenderedDomContextImpl implements RenderedDomContext {
  public pagesContainer: HTMLElement;
  public zoom: number;
  private readonly queries: DisplayListQueries | null;
  private readonly projector: DisplayListPageProjector | null;

  constructor(pagesContainer: HTMLElement, zoom: number = 1, options?: RenderedDomContextOptions) {
    this.pagesContainer = pagesContainer;
    this.zoom = zoom;
    this.queries = options?.displayListQueries ?? null;
    this.projector = options?.projector ?? null;
    if ((this.queries === null) !== (this.projector === null)) {
      throw new Error('RenderedDomContext requires displayListQueries and projector together');
    }
  }

  /**
   * Run spans (`data-doc-start`/`data-doc-end`) in reading order, deep-queried
   * from the a11y mirror subtree nested under `.canvas-pages`.
   */
  private spanEntries(): readonly MirrorSpanEntry[] {
    const out: MirrorSpanEntry[] = [];
    const bodyScopes = Array.from(
      this.pagesContainer.querySelectorAll<HTMLElement>('.layout-page-content')
    );
    const scopes = bodyScopes.length > 0 ? bodyScopes : [this.pagesContainer];
    for (const scope of scopes) {
      for (const el of Array.from(
        scope.querySelectorAll<HTMLElement>('span[data-doc-start][data-doc-end]')
      )) {
        out.push({ el, start: Number(el.dataset.docStart), end: Number(el.dataset.docEnd) });
      }
    }
    return out;
  }

  /** Empty-paragraph runs, deep-queried from the a11y mirror. */
  private emptyRuns(): readonly HTMLElement[] {
    return Array.from(this.pagesContainer.querySelectorAll<HTMLElement>('.layout-empty-run'));
  }

  /**
   * Get pixel coordinates for a display position.
   * Uses the browser's text rendering via Range API for precise positioning.
   */
  getCoordinatesForPosition(position: number): PositionCoordinates | null {
    if (this.queries && this.projector) {
      const rect = this.queries.caretRect(position);
      const projected = rect ? this.projector.projectRect(rect) : null;
      return projected ? { x: projected.x, y: projected.y, height: projected.height } : null;
    }
    const containerRect = this.pagesContainer.getBoundingClientRect();

    // Find spans with display-position data via the mirror's data-doc-* contract
    for (const { el: spanEl, start: spanStart, end: spanEnd } of this.spanEntries()) {
      const span = spanEl;
      // Handle tab spans with exclusive end (tab at [5,6) means pos 6 is next run)
      if (spanEl.classList.contains('layout-run-tab')) {
        if (position >= spanStart && position < spanEnd) {
          const spanRect = spanEl.getBoundingClientRect();
          const lineEl = spanEl.closest('.layout-line');
          const lineHeight = lineEl ? (lineEl as HTMLElement).offsetHeight : 16;

          return {
            x: (spanRect.left - containerRect.left) / this.zoom,
            y: (spanRect.top - containerRect.top) / this.zoom,
            height: lineHeight / this.zoom,
          };
        }
        continue;
      }

      // For text runs, use inclusive range
      if (
        position >= spanStart &&
        position <= spanEnd &&
        span.firstChild?.nodeType === Node.TEXT_NODE
      ) {
        const textNode = span.firstChild as Text;
        const charIndex = Math.min(position - spanStart, textNode.length);

        // Create a range at the exact character position
        const ownerDoc = spanEl.ownerDocument;
        if (!ownerDoc) continue;

        const range = ownerDoc.createRange();
        range.setStart(textNode, charIndex);
        range.setEnd(textNode, charIndex);

        const rangeRect = range.getBoundingClientRect();
        const lineEl = spanEl.closest('.layout-line');
        const lineHeight = lineEl ? (lineEl as HTMLElement).offsetHeight : 16;

        return {
          x: (rangeRect.left - containerRect.left) / this.zoom,
          y: (rangeRect.top - containerRect.top) / this.zoom,
          height: lineHeight / this.zoom,
        };
      }
    }

    // Fallback: try to find position in empty paragraphs
    for (const emptyRun of this.emptyRuns()) {
      const paragraph = emptyRun.closest('.layout-paragraph') as HTMLElement;
      if (!paragraph) continue;

      const paragraphStart = Number(paragraph.dataset.docStart);
      const paragraphEnd = Number(paragraph.dataset.docEnd);

      if (position >= paragraphStart && position <= paragraphEnd) {
        const runRect = emptyRun.getBoundingClientRect();
        const lineEl = emptyRun.closest('.layout-line');
        const lineHeight = lineEl ? (lineEl as HTMLElement).offsetHeight : 16;

        return {
          x: (runRect.left - containerRect.left) / this.zoom,
          y: (runRect.top - containerRect.top) / this.zoom,
          height: lineHeight / this.zoom,
        };
      }
    }

    return null;
  }

  /**
   * Find DOM elements that overlap with a display-position range.
   */
  findElementsForRange(from: number, to: number): Element[] {
    const elements: Element[] = [];
    for (const { el, start, end } of this.spanEntries()) {
      // Check if this span overlaps with the range
      if (end > from && start < to) {
        elements.push(el);
      }
    }
    return elements;
  }

  /**
   * Get bounding rectangles for a range of text.
   * Handles line wraps by returning multiple rects.
   */
  getRectsForRange(
    from: number,
    to: number
  ): Array<{ x: number; y: number; width: number; height: number }> {
    if (this.queries && this.projector) {
      return this.queries
        .rangeRects(from, to)
        .map((rect) => this.projector!.projectRect(rect))
        .filter((rect): rect is ProjectedDisplayListRect => rect !== null);
    }
    const containerRect = this.pagesContainer.getBoundingClientRect();
    const rects: Array<{ x: number; y: number; width: number; height: number }> = [];

    for (const { el: spanEl, start: spanStart, end: spanEnd } of this.spanEntries()) {
      const span = spanEl;
      // Check if this span overlaps with selection
      if (spanEnd > from && spanStart < to) {
        // Handle tab spans - highlight full visual width
        if (spanEl.classList.contains('layout-run-tab')) {
          const spanRect = spanEl.getBoundingClientRect();
          rects.push({
            x: (spanRect.left - containerRect.left) / this.zoom,
            y: (spanRect.top - containerRect.top) / this.zoom,
            width: spanRect.width / this.zoom,
            height: spanRect.height / this.zoom,
          });
          continue;
        }

        if (span.firstChild?.nodeType !== Node.TEXT_NODE) continue;

        const textNode = span.firstChild as Text;
        const ownerDoc = spanEl.ownerDocument;
        if (!ownerDoc) continue;

        // Calculate character range within this span
        const startChar = Math.max(0, from - spanStart);
        const endChar = Math.min(textNode.length, to - spanStart);

        if (startChar < endChar) {
          const range = ownerDoc.createRange();
          range.setStart(textNode, startChar);
          range.setEnd(textNode, endChar);

          // Get all client rects (handles line wraps)
          const clientRects = range.getClientRects();
          for (const rect of Array.from(clientRects)) {
            rects.push({
              x: (rect.left - containerRect.left) / this.zoom,
              y: (rect.top - containerRect.top) / this.zoom,
              width: rect.width / this.zoom,
              height: rect.height / this.zoom,
            });
          }
        }
      }
    }

    return rects;
  }

  getCoordinatesForHfPosition(
    region: 'header' | 'footer',
    rId: string,
    position: number,
    pageIndex: number
  ): PositionCoordinates | null {
    if (!this.queries || !this.projector) return null;
    const rect = this.queries
      .hfCaretRects(region, rId, position)
      .find((candidate) => candidate.pageIndex === pageIndex);
    const projected = rect ? this.projector.projectRect(rect) : null;
    return projected ? { x: projected.x, y: projected.y, height: projected.height } : null;
  }

  getRectsForHfRange(
    region: 'header' | 'footer',
    rId: string,
    from: number,
    to: number,
    pageIndex: number
  ): ProjectedDisplayListRect[] {
    if (!this.queries || !this.projector) return [];
    return this.queries
      .hfRangeRects(region, rId, from, to)
      .filter((rect) => rect.pageIndex === pageIndex)
      .map((rect) => this.projector!.projectRect(rect))
      .filter((rect): rect is ProjectedDisplayListRect => rect !== null);
  }

  getPageBounds(pageIndex: number): ProjectedDisplayListRect | null {
    if (this.projector) return this.projector.getPageBounds(pageIndex);
    const pages = this.pagesContainer.querySelectorAll<HTMLElement>('.layout-page');
    const page = Array.from(pages).find(
      (element, index) => Number(element.dataset.pageIndex ?? index) === pageIndex
    );
    if (!page) return null;
    const pageRect = page.getBoundingClientRect();
    const containerRect = this.pagesContainer.getBoundingClientRect();
    return {
      x: (pageRect.left - containerRect.left) / this.zoom,
      y: (pageRect.top - containerRect.top) / this.zoom,
      width: pageRect.width / this.zoom,
      height: pageRect.height / this.zoom,
    };
  }

  /**
   * Get the offset of the pages container from its parent viewport.
   * This is needed for positioning overlays that are rendered in the
   * viewport container rather than directly in the pages container.
   */
  getContainerOffset(): { x: number; y: number } {
    const parent = this.pagesContainer.parentElement;
    if (!parent) return { x: 0, y: 0 };

    const containerRect = this.pagesContainer.getBoundingClientRect();
    const parentRect = parent.getBoundingClientRect();

    return {
      x: (containerRect.left - parentRect.left) / this.zoom,
      y: (containerRect.top - parentRect.top) / this.zoom,
    };
  }
}

/**
 * Create a RenderedDomContext for a pages container element.
 *
 * @param pagesContainer - The `.canvas-pages` host holding the rendered pages
 * @param zoom - Current zoom level (default 1)
 * @param options - query/projector backend options. Omit for the mirror-DOM
 *   fallback.
 */
export function createRenderedDomContext(
  pagesContainer: HTMLElement,
  zoom: number = 1,
  options?: RenderedDomContextOptions
): RenderedDomContext {
  return new RenderedDomContextImpl(pagesContainer, zoom, options);
}

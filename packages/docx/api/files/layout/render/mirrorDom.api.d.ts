// @betteroffice/docx declaration layout/render/mirrorDom.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Builds an accessible display-list mirror with DOM APIs only. */
import type { DisplayPage } from './displayList';
export declare const MIRROR_CLASS_NAMES: {
    page: string;
    content: string;
    header: string;
    footer: string;
    block: string;
    table: string;
    tableRow: string;
    tableCell: string;
    text: string;
    rect: string;
    image: string;
    shape: string;
    chart: string;
    chartTable: string;
    chartRow: string;
    chartCell: string;
    decoration: string;
    line: string;
    pageBorder: string;
    tableCut: string;
    blockSdt: string;
    inlineSdtWidget: string;
    revisionPmark: string;
    revisionPmarkGlyph: string;
    revisionPmarkIns: string;
    revisionPmarkDel: string;
    listMarker: string;
    revisionTable: string;
    revisionRow: string;
    revisionCell: string;
    revisionIns: string;
    revisionDel: string;
    revisionMerge: string;
    notes: string;
    note: string;
    noteBacklink: string;
};
/**
 * screen-reader labels for the mirror's structural wrappers. all strings are
 * user-facing — hosts pass localized values (React adapter: `t('a11y.*')`);
 * core deliberately ships no English defaults so no locale bypasses i18n.
 */
export interface MirrorLabels {
    /** page container label, e.g. "Page 3" */
    page?: string;
    /** header band label, e.g. "Page header" */
    header?: string;
    /** footer band label, e.g. "Page footer" */
    footer?: string;
    /**
     * note→reference backlink label, e.g. "Back to reference". Optional: without
     * it the backlink anchor falls back to the note's own label (a number, not
     * hardcoded English), so no locale bypasses i18n.
     */
    noteBacklink?: string;
    /** description of a chart that caches no points, e.g. "No data"; without
     *  it such a chart gets no description */
    chartNoData?: string;
    /** name of a picture without alternative text, e.g. "Picture"; without it such a picture gets no image role */
    picture?: string;
    /** word a visible list bullet reads as, e.g. "bullet"; without it a bullet stays hidden */
    bullet?: string;
}
export interface BuildMirrorPageOptions {
    /** document to create elements in (default: global document) */
    document?: Document;
    /** localized aria-labels for page/header/footer wrappers */
    labels?: MirrorLabels;
}
/** Whether `buildMirrorPage(page)` may hold a table header cell, which cells on other pages may name; a cell marked as a header with nothing to read reads as a plain cell. */
export declare function mirrorPageHasHeaderCells(page: DisplayPage): boolean;
/** Whether `buildMirrorPage(page)` holds a link Tab stops at. */
export declare function mirrorPageHasTabStops(page: DisplayPage): boolean;
/**
 * Reduces a `buildMirrorPage` tree, in place, to its links, its notes (the
 * targets of note references), its header cells (which cells on any page may
 * name) and the elements that label what stays. What stays inside a table
 * cell keeps the whole outermost cell, so the cell keeps its accessible name.
 * Every other element and text goes; the ancestors of what stays keep their
 * attributes, so the ids, order, positions and table semantics that remain
 * are the full mirror's.
 */
export declare function reduceMirrorToLinks(mirror: HTMLElement): HTMLElement;
/** `buildMirrorPage(page)` reduced to its links: see `reduceMirrorToLinks`. */
export declare function buildMirrorPageLinks(page: DisplayPage, options?: BuildMirrorPageOptions): HTMLElement;
/**
 * Reduces a `buildMirrorPage` tree, in place, to what a screen reader reads:
 * everything `reduceMirrorToLinks` keeps, unchanged, plus the text of every
 * other run as plain text inside its paragraph, table and region wrappers,
 * and images with an accessible name. Wrappers that hold no kept element lose
 * their ids and data attributes, so painter-contract queries never match them.
 */
export declare function reduceMirrorToText(mirror: HTMLElement): HTMLElement;
/** `buildMirrorPage(page)` reduced to readable text: see `reduceMirrorToText`. */
export declare function buildMirrorPageText(page: DisplayPage, options?: BuildMirrorPageOptions): HTMLElement;
/** Whether `page`'s mirror holds the note or note reference with element id `id`. */
export declare function displayPageHoldsMirrorId(page: DisplayPage, id: string): boolean;
/**
 * build the mirror DOM for one display page: an invisible (opacity 0,
 * pointer-events none) absolutely-positioned tree whose elements sit at the
 * exact page-local coordinates of their display primitives, carry the painter
 * dataset contract, and expose real text nodes in reading order for
 * accessibility. primitives sharing a blockId group under one block wrapper
 * (role=paragraph, or role=table with row/cell structure when primitives
 * carry `cell` grid coordinates) stamped with the union of the children's doc
 * range; blockless primitives (table borders, separators) attach directly to
 * `.layout-page-content`.
 *
 * header/footer regions mirror the painter's structural classes: each HfRegion
 * becomes a `.layout-page-header` / `.layout-page-footer` wrapper at the band's
 * page-local y, and all of the region's primitives (including their block
 * wrappers and doc-range attrs, which refer to the HF doc identified by
 * `data-hf-rid`) stay scoped inside it — never under `.layout-page-content`.
 *
 * data-from-line/to-line ride on paragraph fragment primitives (body, HF, text
 * box) and land on the paragraph wrapper; data-vmerge-continuation comes from
 * the cell's `continuation` flag and lands on the ARIA cell. still not derivable
 * from the v0 display list and therefore not emitted: data-carried-from-prev/
 * to-next.
 */
export declare function buildMirrorPage(page: DisplayPage, options?: BuildMirrorPageOptions): HTMLElement;
/**
 * Brings `mirror`, a `buildMirrorPage` tree, up to date with `page` in place,
 * exactly as `buildMirrorPage(page, options)` would build it. Blocks whose
 * content is unchanged keep their elements; when only their document
 * positions moved, those attributes shift in place. Returns false, changing
 * nothing, when `mirror` is not a tree it can update.
 */
export declare function updateMirrorPage(mirror: HTMLElement, page: DisplayPage, options?: BuildMirrorPageOptions): boolean;

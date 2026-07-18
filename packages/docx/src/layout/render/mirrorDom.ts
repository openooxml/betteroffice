// accessibility mirror: a geometry-faithful DOM tree generated from a
// DisplayPage. it speaks the stable mirror DOM contract (data-doc-start/end,
// data-block-id, data-comment-id, revision attrs, .layout-page-content,
// .layout-run-text, .layout-table-cut-border) so selection mapping, comment
// anchors, and the Playwright suite resolve against it under the canvas
// renderer.
//
// it is also the accessible content of the canvas renderer (the design's
// phase-1 gate): real text nodes in reading order, paragraph/table ARIA
// semantics, tracked-change roles. the tree is invisible (opacity 0,
// pointer-events none) but deliberately NOT aria-hidden — under a canvas
// there is nothing else for a screen reader to read.
//
// security: every value here is file-derived. nodes are built exclusively via
// createElement/setAttribute/textContent — never innerHTML (see the repo security guidelines
// "Security — untrusted DOCX/HTML input").

import type {
  DocAttrs,
  DisplayPage,
  DisplayPrimitive,
  DisplayTableMetadata,
  HfRegion,
  NoteRegion,
  NoteRegionNote,
  SdtAttrs,
  StructuralRevision,
  TableCellRef,
} from './displayList';
import { textRunRect, glyphRunRect, lineRect, type GeoRect } from './displayListGeometry';

export const MIRROR_CLASS_NAMES = {
  page: 'layout-page layout-page-mirror',
  content: 'layout-page-content',
  header: 'layout-page-header',
  footer: 'layout-page-footer',
  block: 'layout-paragraph',
  // table wrappers reuse the painter's structural classes (renderTable's
  // TABLE_CLASS_NAMES) so contract consumers see one vocabulary
  table: 'layout-table',
  tableRow: 'layout-table-row',
  tableCell: 'layout-table-cell',
  text: 'layout-run layout-run-text',
  rect: 'layout-mirror-rect',
  image: 'layout-run layout-run-image',
  shape: 'layout-run layout-run-shape',
  chart: 'layout-run layout-run-chart',
  decoration: 'layout-decoration',
  line: 'layout-mirror-line',
  pageBorder: 'layout-page-border',
  tableCut: 'layout-table-cut-border',
  blockSdt: 'layout-block-sdt',
  inlineSdtWidget: 'layout-inline-sdt-widget',
  revisionPmark: 'layout-revision-pmark',
  revisionPmarkGlyph: 'layout-revision-pmark-glyph',
  revisionPmarkIns: 'layout-revision-ins',
  revisionPmarkDel: 'layout-revision-del',
  listMarker: 'layout-list-marker',
  revisionTable: 'oox-revision-table',
  revisionRow: 'oox-revision-row',
  revisionCell: 'oox-revision-cell',
  revisionIns: 'oox-revision-ins',
  revisionDel: 'oox-revision-del',
  revisionMerge: 'oox-revision-merge',
  notes: 'layout-page-notes',
  note: 'layout-note',
  noteBacklink: 'layout-note-backlink',
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
}

/** per-page build state threaded through the mirror walk */
interface MirrorBuildCtx {
  labels?: MirrorLabels;
  /**
   * note-ref anchor ids already assigned on this page: a reference mark split
   * into several primitives (bidi/font subranges) must yield exactly one
   * element id for the note's doc-backlink to target.
   */
  noteRefIds: Set<string>;
}

export interface BuildMirrorPageOptions {
  /** document to create elements in (default: global document) */
  document?: Document;
  /** localized aria-labels for page/header/footer wrappers */
  labels?: MirrorLabels;
}

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
 * wrappers and doc-range attrs, which refer to the HF PM doc identified by
 * `data-hf-rid`) stay scoped inside it — never under `.layout-page-content`.
 *
 * data-from-line/to-line ride on paragraph fragment primitives (body, HF, text
 * box) and land on the paragraph wrapper; data-vmerge-continuation comes from
 * the cell's `continuation` flag and lands on the ARIA cell. still not derivable
 * from the v0 display list and therefore not emitted: data-carried-from-prev/
 * to-next (display-list contract gap — see openspec/changes/rust-canvas-engine).
 */
export function buildMirrorPage(
  page: DisplayPage,
  options: BuildMirrorPageOptions = {}
): HTMLElement {
  const doc = options.document ?? document;
  const labels = options.labels;
  const ctx: MirrorBuildCtx = { labels, noteRefIds: new Set() };

  const pageEl = doc.createElement('div');
  pageEl.className = MIRROR_CLASS_NAMES.page;
  pageEl.setAttribute('role', 'document');
  if (labels?.page) pageEl.setAttribute('aria-label', labels.page);
  pageEl.dataset.pageIndex = String(page.pageIndex);
  if (page.sectionId !== undefined) pageEl.dataset.sectionId = page.sectionId;
  if (page.sectionIndex !== undefined) pageEl.dataset.sectionIndex = String(page.sectionIndex);
  if (page.sectionPageIndex !== undefined) {
    pageEl.dataset.sectionPageIndex = String(page.sectionPageIndex);
  }
  if (page.sectionPageNumber !== undefined) {
    pageEl.dataset.sectionPageNumber = String(page.sectionPageNumber);
  }
  if (page.pageLabel !== undefined) pageEl.dataset.pageLabel = page.pageLabel;
  pageEl.style.position = 'relative';
  pageEl.style.width = `${page.width}px`;
  pageEl.style.height = `${page.height}px`;
  pageEl.style.opacity = '0';
  pageEl.style.pointerEvents = 'none';

  const contentEl = doc.createElement('div');
  contentEl.className = MIRROR_CLASS_NAMES.content;
  contentEl.style.position = 'absolute';
  contentEl.style.left = '0';
  contentEl.style.top = '0';
  contentEl.style.width = `${page.width}px`;
  contentEl.style.height = `${page.height}px`;
  pageEl.appendChild(contentEl);
  for (const border of (page.pageBorders ?? []).filter((p) => p.zOrder === 'back')) {
    contentEl.appendChild(renderPageBorderMirror(border, doc, 0));
  }
  appendMirrorPrimitives(contentEl, page.primitives, doc, 0, ctx);
  for (const noteArea of page.noteAreas ?? []) {
    contentEl.appendChild(buildMirrorNoteArea(noteArea, page, doc, ctx));
  }

  // painter append order: content, then header, then footer
  for (const region of [page.header, page.footer]) {
    if (region) pageEl.appendChild(buildMirrorRegion(region, page, doc, ctx));
  }
  for (const border of (page.pageBorders ?? []).filter((p) => p.zOrder !== 'back')) {
    pageEl.appendChild(renderPageBorderMirror(border, doc, 0));
  }

  return pageEl;
}

// one HfRegion → its painter-contract wrapper. children are placed with a
// -region.y offset so the display list's page coordinates land at the same
// page-local position once the wrapper's own top is added back (the harvest
// walk accumulates ancestor offsets).
function buildMirrorRegion(
  region: HfRegion,
  page: DisplayPage,
  doc: Document,
  ctx: MirrorBuildCtx
): HTMLElement {
  const regionEl = doc.createElement('div');
  regionEl.className =
    region.kind === 'header' ? MIRROR_CLASS_NAMES.header : MIRROR_CLASS_NAMES.footer;
  regionEl.dataset.hfRid = region.rId;
  // role=region needs an accessible name to be exposed as a landmark, so the
  // role rides with the label (host-localized); unlabeled bands stay generic
  const label = region.kind === 'header' ? ctx.labels?.header : ctx.labels?.footer;
  if (label) {
    regionEl.setAttribute('role', 'region');
    regionEl.setAttribute('aria-label', label);
  }
  regionEl.style.position = 'absolute';
  regionEl.style.left = '0';
  regionEl.style.top = `${region.y}px`;
  regionEl.style.width = `${page.width}px`;
  regionEl.style.height = `${region.height}px`;
  appendMirrorPrimitives(regionEl, region.primitives, doc, -region.y, ctx);
  return regionEl;
}

// primitives of one block, collected before the wrapper is populated so table
// blocks can be restructured into row/cell wrappers
interface BlockGroup {
  el: HTMLElement;
  prims: Exclude<DisplayPrimitive, { kind: 'line' }>[];
  table?: DisplayTableMetadata;
  cells?: Array<{ el: HTMLElement; bounds?: GeoRect }>;
}

// shared body/region flow: primitives sharing a blockId group under one
// block wrapper (created in first-seen paint order so text reads in document
// order), each wrapper stamped with the union doc range of its children so it
// satisfies the fragment-level part of the painter contract. blockIds (and
// doc ranges) are container-scoped — a header's block 0 is a different doc
// than the body's block 0, which is why the maps never cross containers.
//
// table primitives group by their stamped OWNING table fragment instead of
// their source block: the builder emits cell paragraphs under the cell
// paragraph's own block id, but the painter rendered exactly one
// `.layout-table` per table fragment — grouping per cell paragraph produced
// one role=table PER CELL (screen readers announced a 1×2 header table as two
// tables; `.layout-page-header [data-row-start]` matched twice). The fragment
// key includes rowStart so a table split across columns on one page keeps one
// wrapper per fragment, matching the painter. Nested tables carry their own
// tableId (the outer stamp only adds parentTableId), so they keep their own
// group and reparent through nestTableBlocks.
function appendMirrorPrimitives(
  container: HTMLElement,
  primitives: DisplayPrimitive[],
  doc: Document,
  offsetY: number,
  ctx: MirrorBuildCtx
): void {
  const blocks = new Map<number | string, BlockGroup>();

  for (const p of primitives) {
    // live-pipeline blocks carry string ids in blockKey, goldens numeric ids
    // in blockId — exactly one is set when the primitive has block identity
    const blockId = p.kind === 'line' ? undefined : (p.blockKey ?? p.blockId);
    if (blockId === undefined) {
      container.appendChild(renderMirrorPrimitive(p, doc, offsetY, ctx));
      continue;
    }
    const table = p.table;
    const key =
      table?.tableId !== undefined ? `table:${table.tableId}#${table.rowStart ?? 0}` : blockId;
    let group = blocks.get(key);
    if (!group) {
      const blockEl = doc.createElement('div');
      // a merged table group answers for the table block itself — the painter
      // contract's data-block-id names the table block, not a cell paragraph
      blockEl.dataset.blockId = String(table?.tableId ?? blockId);
      blockEl.style.position = 'absolute';
      blockEl.style.left = '0';
      blockEl.style.top = '0';
      container.appendChild(blockEl);
      group = { el: blockEl, prims: [] };
      blocks.set(key, group);
    }
    group.prims.push(p as BlockGroup['prims'][number]);
  }

  for (const group of blocks.values()) {
    populateBlock(group, doc, offsetY, ctx);
  }
  nestTableBlocks(blocks);
}

// finalize one block wrapper: paragraph blocks keep flat paint order; blocks
// whose primitives carry `cell` grid coordinates become an ARIA table (rows
// ascending, cells by column — table reading order regardless of paint order)
function populateBlock(
  group: BlockGroup,
  doc: Document,
  offsetY: number,
  ctx: MirrorBuildCtx
): void {
  const { el, prims } = group;
  const isTable = prims.some(
    (p) =>
      p.cell !== undefined ||
      p.structuralRevision?.scope === 'table' ||
      p.structuralRevision?.scope === 'row' ||
      p.structuralRevision?.scope === 'cell'
  );

  if (!isTable) {
    populateParagraphBlock(el, prims, doc, offsetY, ctx);
    return;
  }

  el.className = MIRROR_CLASS_NAMES.table;
  el.setAttribute('role', 'table');
  const table = prims.find((p) => p.table !== undefined)?.table;
  group.table = table;
  group.cells = [];
  if (table?.tableId !== undefined) el.dataset.tableId = table.tableId;
  if (table?.rowStart !== undefined) el.dataset.rowStart = String(table.rowStart);
  if (table?.rowEnd !== undefined) el.dataset.rowEnd = String(table.rowEnd);
  if (table?.headerRowCount !== undefined) {
    el.dataset.headerRowCount = String(table.headerRowCount);
  }
  if (table?.parentTableId !== undefined) el.dataset.parentTableId = table.parentTableId;
  if (table?.caption) el.setAttribute('aria-label', table.caption);
  if (table?.description) el.setAttribute('aria-description', table.description);
  const tableRevision = findStructuralRevision(prims, 'table');
  if (tableRevision) applyStructuralRevisionAttrs(el, tableRevision);
  applyBlockSdtAttrs(el, prims.find((p) => p.sdt !== undefined)?.sdt);

  // cell-less primitives of a table block (shading rects, decorations the
  // builder didn't attribute to a cell) stay direct children, ahead of the
  // row flow so cells read contiguously
  for (const p of prims) {
    if (p.cell === undefined) el.appendChild(renderMirrorPrimitive(p, doc, offsetY, ctx));
  }

  // group by (row, col) preserving paint order inside a cell
  const rows = new Map<number, Map<number, { ref: TableCellRef; prims: BlockGroup['prims'] }>>();
  let rowCount = 0;
  let colCount = 0;
  for (const p of prims) {
    if (p.cell === undefined) continue;
    const { row, col } = p.cell;
    rowCount = Math.max(rowCount, row + (p.cell.rowSpan ?? 1));
    colCount = Math.max(colCount, col + (p.cell.colSpan ?? 1));
    let cells = rows.get(row);
    if (!cells) {
      cells = new Map();
      rows.set(row, cells);
    }
    let cell = cells.get(col);
    if (!cell) {
      cell = { ref: p.cell, prims: [] };
      cells.set(col, cell);
    }
    cell.prims.push(p);
  }
  el.setAttribute('aria-rowcount', String(table?.rowCount ?? rowCount));
  el.setAttribute('aria-colcount', String(table?.columnCount ?? colCount));

  for (const rowIndex of [...rows.keys()].sort((a, b) => a - b)) {
    const rowEl = doc.createElement('div');
    rowEl.className = MIRROR_CLASS_NAMES.tableRow;
    rowEl.setAttribute('role', 'row');
    rowEl.setAttribute('aria-rowindex', String(rowIndex + 1));
    const rowRevision = findStructuralRevision(prims, 'row', rowIndex);
    if (rowRevision) applyStructuralRevisionAttrs(rowEl, rowRevision);
    rowEl.style.position = 'absolute';
    rowEl.style.left = '0';
    rowEl.style.top = '0';
    el.appendChild(rowEl);

    const cells = rows.get(rowIndex)!;
    for (const colIndex of [...cells.keys()].sort((a, b) => a - b)) {
      const cell = cells.get(colIndex)!;
      const cellEl = doc.createElement('div');
      cellEl.className = MIRROR_CLASS_NAMES.tableCell;
      cellEl.setAttribute('role', cell.ref.isHeader ? 'columnheader' : 'cell');
      cellEl.setAttribute('aria-colindex', String(colIndex + 1));
      if (cell.ref.cellId) cellEl.id = cell.ref.cellId;
      if (cell.ref.headerIds?.length) {
        cellEl.setAttribute('aria-labelledby', cell.ref.headerIds.join(' '));
      }
      if (cell.ref.repeatedHeader) cellEl.dataset.repeatedHeader = 'true';
      if (cell.ref.noWrap) cellEl.dataset.noWrap = 'true';
      const rowSpan = cell.ref.rowSpan ?? 1;
      const colSpan = cell.ref.colSpan ?? 1;
      if (rowSpan > 1) cellEl.setAttribute('aria-rowspan', String(rowSpan));
      if (colSpan > 1) cellEl.setAttribute('aria-colspan', String(colSpan));
      const cellRevision = findStructuralRevision(cell.prims, 'cell', rowIndex, colIndex);
      if (cellRevision) applyStructuralRevisionAttrs(cellEl, cellRevision);
      // synthetic slice of a vertically-merged cell re-painted on a continuation
      // page → data-vmerge-continuation (painter contract); not selectable, so
      // the builder already stripped its doc positions.
      if (cell.ref.continuation) cellEl.dataset.vmergeContinuation = 'true';
      cellEl.style.position = 'absolute';
      cellEl.style.left = '0';
      cellEl.style.top = '0';
      // One Word table owns one ARIA table wrapper, but its cell paragraphs
      // remain separate painter-contract blocks. Flattening every primitive
      // directly into the cell loses the paragraph node boundary: an empty
      // cell then exposes only its zero-width caret marker at pmStart+1, and
      // structural consumers advance a second time into the next cell. Keep
      // the source block identity and paragraph range beneath the merged table.
      const cellBlocks = new Map<number | string, BlockGroup['prims']>();
      const blockless: BlockGroup['prims'] = [];
      for (const p of cell.prims) {
        const blockId = p.blockKey ?? p.blockId;
        if (blockId === undefined) {
          blockless.push(p);
          continue;
        }
        const blockPrims = cellBlocks.get(blockId);
        if (blockPrims) blockPrims.push(p);
        else cellBlocks.set(blockId, [p]);
      }
      for (const p of inLogicalOrder(blockless)) {
        cellEl.appendChild(renderMirrorPrimitive(p, doc, offsetY, ctx));
      }
      for (const [blockId, blockPrims] of cellBlocks) {
        const paragraphEl = doc.createElement('div');
        paragraphEl.dataset.blockId = String(blockId);
        paragraphEl.style.position = 'absolute';
        paragraphEl.style.left = '0';
        paragraphEl.style.top = '0';
        populateParagraphBlock(paragraphEl, blockPrims, doc, offsetY, ctx);
        cellEl.appendChild(paragraphEl);
      }
      stampDocRangeFromDescendants(cellEl, cell.prims);
      applyContainerLanguageAndDirection(cellEl, cell.prims);
      group.cells.push({ el: cellEl, bounds: boundsOfPrimitives(cell.prims) });
      rowEl.appendChild(cellEl);
    }
  }

  stampDocRange(el, prims);
}

/** Populate one paragraph-like block without re-interpreting table metadata. */
function populateParagraphBlock(
  el: HTMLElement,
  prims: BlockGroup['prims'],
  doc: Document,
  offsetY: number,
  ctx: MirrorBuildCtx
): void {
  const chart = prims.find((p) => p.chart !== undefined)?.chart;
  el.className = chart
    ? `${MIRROR_CLASS_NAMES.block} ${MIRROR_CLASS_NAMES.chart}`
    : MIRROR_CLASS_NAMES.block;
  el.setAttribute('role', chart ? 'img' : 'paragraph');
  if (chart?.label) el.setAttribute('aria-label', chart.label);
  for (const p of inLogicalOrder(prims)) {
    el.appendChild(renderMirrorPrimitive(p, doc, offsetY, ctx));
  }
  const pmarkRevision = findStructuralRevision(prims, 'pmark');
  if (pmarkRevision) applyStructuralRevisionAttrs(el, pmarkRevision);
  stampParagraphDocRange(el, prims, doc, offsetY);
  applyBlockSdtAttrs(el, prims.find((p) => p.sdt !== undefined)?.sdt);
  const paraId = prims.find((p) => p.paraId !== undefined)?.paraId;
  if (paraId !== undefined) el.dataset.paraId = paraId;
  const withLineRange = prims.find((p) => p.fromLine !== undefined);
  if (withLineRange?.fromLine !== undefined) {
    el.dataset.fromLine = String(withLineRange.fromLine);
    if (withLineRange.toLine !== undefined) el.dataset.toLine = String(withLineRange.toLine);
  }
  applyContainerLanguageAndDirection(el, prims);
  applyBlockObjectA11y(el, prims);
}

/**
 * Paragraph wrappers describe the structural node, while a zero-width empty
 * run describes its valid caret position inside that node. Rust emits the
 * latter at pmStart+1; retain that exact value on the run, but recover pmStart
 * on the wrapper so scrolling and cell-selection consumers can distinguish
 * the paragraph boundary from its content position.
 */
function stampParagraphDocRange(
  el: HTMLElement,
  prims: BlockGroup['prims'],
  doc: Document,
  offsetY: number
): void {
  stampDocRange(el, prims);
  const stampedStart = Number(el.dataset.docStart);
  const stampedEnd = Number(el.dataset.docEnd);
  const fragmentStarts = prims
    .map((p) => p.fragmentDocStart)
    .filter((value): value is number => Number.isFinite(value));
  const fragmentEnds = prims
    .map((p) => p.fragmentDocEnd)
    .filter((value): value is number => Number.isFinite(value));
  const fragmentStart = fragmentStarts.length > 0 ? Math.min(...fragmentStarts) : undefined;
  const fragmentEnd = fragmentEnds.length > 0 ? Math.max(...fragmentEnds) : undefined;
  // Compatibility for old display-list fixtures: a lone zero-width marker is
  // known to be an empty paragraph's content position at pmStart+1.
  const structuralStart =
    fragmentStart ??
    (Number.isFinite(stampedStart) && stampedStart === stampedEnd && stampedStart > 0
      ? stampedStart - 1
      : undefined);
  if (structuralStart !== undefined) el.dataset.docStart = String(structuralStart);
  if (fragmentEnd !== undefined) el.dataset.docEnd = String(fragmentEnd);

  // The wrapper stays at (0,0) so its absolutely positioned children retain
  // page-local geometry. Give a fragment start that has no inline primitive
  // its own zero-width, line-height-tall anchor at the first primitive's exact
  // coordinates; viewport consumers can then find pmStart without stealing
  // the selectable run/caret marker's position.
  if (structuralStart !== undefined && !prims.some((p) => p.docStart === structuralStart)) {
    const firstPositioned = inLogicalOrder(prims).find((p) => primitiveRect(p) !== undefined);
    const rect = firstPositioned ? primitiveRect(firstPositioned) : undefined;
    if (rect) {
      const anchor = doc.createElement('span');
      anchor.className = 'layout-mirror-structural-anchor';
      anchor.dataset.docStart = String(structuralStart);
      anchor.dataset.docEnd = String(structuralStart);
      placeAt(anchor, rect, offsetY);
      el.insertBefore(anchor, el.firstChild);
    }
  }
}

/** Union paragraph structural ranges, falling back to primitive ranges. */
function stampDocRangeFromDescendants(el: HTMLElement, fallbackPrims: BlockGroup['prims']): void {
  const descendants = el.querySelectorAll<HTMLElement>('.layout-paragraph[data-doc-start]');
  let start = Infinity;
  let end = -Infinity;
  for (const descendant of descendants) {
    const descendantStart = Number(descendant.dataset.docStart);
    const descendantEnd = Number(descendant.dataset.docEnd);
    if (Number.isFinite(descendantStart)) start = Math.min(start, descendantStart);
    if (Number.isFinite(descendantEnd)) end = Math.max(end, descendantEnd);
  }
  if (start !== Infinity) el.dataset.docStart = String(start);
  if (end !== -Infinity) el.dataset.docEnd = String(end);
  if (start === Infinity && end === -Infinity) stampDocRange(el, fallbackPrims);
}

/**
 * Move a nested table below the geometrically owning outer cell. Batch F emits
 * the parent table id and per-cell clip rectangles, so the mirror can recover
 * hierarchy without changing visual coordinates. Legacy lists without clips
 * remain flat rather than guessing the wrong owner.
 */
function nestTableBlocks(blocks: Map<number | string, BlockGroup>): void {
  const byTableId = new Map<string, BlockGroup>();
  for (const group of blocks.values()) {
    if (group.table?.tableId) byTableId.set(group.table.tableId, group);
  }
  for (const group of blocks.values()) {
    const parentId = group.table?.parentTableId;
    if (!parentId) continue;
    const parent = byTableId.get(parentId);
    const childBounds = boundsOfPrimitives(group.prims);
    if (!parent?.cells?.length || !childBounds) continue;
    const centerX = childBounds.x + childBounds.w / 2;
    const centerY = childBounds.y + childBounds.h / 2;
    const candidates = parent.cells.filter(
      (cell) => cell.bounds && pointInRect(centerX, centerY, cell.bounds)
    );
    candidates.sort((a, b) => rectArea(a.bounds!) - rectArea(b.bounds!));
    candidates[0]?.el.appendChild(group.el);
  }
}

function pointInRect(x: number, y: number, rect: GeoRect): boolean {
  return x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h;
}

function rectArea(rect: GeoRect): number {
  return rect.w * rect.h;
}

function boundsOfPrimitives(prims: BlockGroup['prims']): GeoRect | undefined {
  let out: GeoRect | undefined;
  for (const p of prims) {
    const clip = p.clipGroup?.clip;
    const rect =
      clip?.x !== undefined && clip.y !== undefined && clip.w !== undefined && clip.h !== undefined
        ? { x: clip.x, y: clip.y, w: clip.w, h: clip.h }
        : primitiveRect(p);
    if (!rect) continue;
    if (!out) {
      out = { ...rect };
      continue;
    }
    const right = Math.max(out.x + out.w, rect.x + rect.w);
    const bottom = Math.max(out.y + out.h, rect.y + rect.h);
    out.x = Math.min(out.x, rect.x);
    out.y = Math.min(out.y, rect.y);
    out.w = right - out.x;
    out.h = bottom - out.y;
  }
  return out;
}

function primitiveRect(p: BlockGroup['prims'][number]): GeoRect | undefined {
  switch (p.kind) {
    case 'text':
      return textRunRect(p);
    case 'glyphRun':
      return glyphRunRect(p);
    case 'rect':
      return { x: p.x, y: p.y, w: p.w, h: p.h };
    case 'image':
    case 'shape':
      return { x: p.x, y: p.y, w: p.w, h: p.h };
    case 'decoration':
      return { x: p.x, y: p.y, w: p.w, h: p.h };
  }
}

/** Reorder only explicitly indexed primitives, leaving paint-only nodes in place. */
function inLogicalOrder<T extends BlockGroup['prims']>(prims: T): T {
  const ordered = prims
    .filter((p) => p.logicalOrder !== undefined)
    .sort((a, b) => {
      const delta = (a.logicalOrder ?? 0) - (b.logicalOrder ?? 0);
      return delta || (a.docStart ?? 0) - (b.docStart ?? 0);
    });
  if (ordered.length < 2) return prims;
  let index = 0;
  return prims.map((p) => (p.logicalOrder === undefined ? p : ordered[index++])) as T;
}

function applyContainerLanguageAndDirection(el: HTMLElement, prims: BlockGroup['prims']): void {
  const lang = prims.find((p) => p.lang)?.lang;
  if (lang) el.lang = lang;
  const directional = prims.find((p) => p.bidiLevel !== undefined || 'rtl' in p);
  if (!directional) return;
  const rtl =
    directional.bidiLevel !== undefined
      ? directional.bidiLevel % 2 === 1
      : 'rtl' in directional
        ? directional.rtl
        : false;
  el.dir = rtl ? 'rtl' : 'ltr';
}

function applyBlockObjectA11y(el: HTMLElement, prims: BlockGroup['prims']): void {
  const attrs = prims.find(
    (p) => p.ariaLabel || p.ariaDescription || p.decorative || p.hiddenObject || p.groupId
  );
  if (!attrs) return;
  applyObjectA11y(el, attrs);
}

// union doc range of a primitive set → data-doc-start/end on the wrapper
// (the fragment-level part of the painter contract)
function stampDocRange(el: HTMLElement, prims: BlockGroup['prims']): void {
  let docStart: number | undefined;
  let docEnd: number | undefined;
  for (const p of prims) {
    if (p.docStart !== undefined) {
      docStart = docStart === undefined ? p.docStart : Math.min(docStart, p.docStart);
    }
    if (p.docEnd !== undefined) {
      docEnd = docEnd === undefined ? p.docEnd : Math.max(docEnd, p.docEnd);
    }
  }
  if (docStart !== undefined) el.dataset.docStart = String(docStart);
  if (docEnd !== undefined) el.dataset.docEnd = String(docEnd);
}

function findStructuralRevision(
  prims: BlockGroup['prims'],
  scope: StructuralRevision['scope'],
  rowIndex?: number,
  colIndex?: number
): StructuralRevision | undefined {
  return prims.find((p) => {
    const rev = p.structuralRevision;
    if (!rev || rev.scope !== scope) return false;
    if (rowIndex !== undefined && rev.rowIndex !== rowIndex) return false;
    if (colIndex !== undefined && rev.colIndex !== colIndex) return false;
    return true;
  })?.structuralRevision;
}

function renderMirrorPrimitive(
  p: DisplayPrimitive,
  doc: Document,
  offsetY: number,
  ctx: MirrorBuildCtx
): HTMLElement {
  switch (p.kind) {
    case 'text': {
      const el = createTextMirrorElement(doc, p, ctx);
      el.className = MIRROR_CLASS_NAMES.text;
      el.textContent = p.text;
      placeAt(el, textRunRect(p), offsetY);
      // single-property CSSOM assignment: a hostile font string cannot escape
      // into other declarations, and geometry never derives from this value
      el.style.font = p.font;
      applyRunLanguageAndDirection(el, p);
      applyTextVisualStyles(el, p);
      applyPrimitiveVisualStyle(
        el,
        p.hidden && p.opacity === undefined ? 0.4 : p.opacity,
        p.rotationDeg,
        p.horizontalScale
      );
      if (p.structuralRevision?.scope === 'pmark') {
        el.classList.add(MIRROR_CLASS_NAMES.revisionPmarkGlyph);
        if (p.structuralRevision.kind === 'ins') {
          el.classList.add(MIRROR_CLASS_NAMES.revisionPmarkIns);
        }
        if (p.structuralRevision.kind === 'del') {
          el.classList.add(MIRROR_CLASS_NAMES.revisionPmarkDel);
        }
        el.setAttribute('aria-hidden', 'true');
      }
      if (p.listMarker) {
        el.classList.add(MIRROR_CLASS_NAMES.listMarker);
        el.setAttribute('aria-hidden', 'true');
        if (p.listMarkerRevision) {
          el.style.color =
            p.listMarkerRevision === 'ins' ? 'rgb(46, 125, 50)' : 'rgb(198, 40, 40)';
        }
      }
      applyDocAttrs(el, p);
      return el;
    }
    case 'glyphRun': {
      // a shaped glyph run mirrors as the same real text node a browser-shaped
      // run does: the run's `text` in reading order, placed at the glyph-derived
      // rect, carrying the painter dataset contract. the canvas paints outlines;
      // the mirror is the accessible/selection surface.
      const el = createTextMirrorElement(doc, p, ctx);
      el.className = MIRROR_CLASS_NAMES.text;
      el.textContent = p.text;
      placeAt(el, glyphRunRect(p), offsetY);
      // numeric size only — no file-derived family string reaches the CSSOM
      el.style.fontSize = `${p.size}px`;
      applyGlyphRunWeightAndStyle(el, p.fallbackFont);
      applyRunLanguageAndDirection(el, p);
      applyTextVisualStyles(el, p);
      applyPrimitiveVisualStyle(
        el,
        p.hidden && p.opacity === undefined ? 0.4 : p.opacity,
        p.rotationDeg,
        p.horizontalScale
      );
      applyDocAttrs(el, p);
      return el;
    }
    case 'rect': {
      const el = doc.createElement('div');
      el.className = MIRROR_CLASS_NAMES.rect;
      placeAt(el, { x: p.x, y: p.y, w: p.w, h: p.h }, offsetY);
      applyDocAttrs(el, p);
      return el;
    }
    case 'line': {
      const el = doc.createElement('div');
      el.className = p.role === 'table-cut' ? MIRROR_CLASS_NAMES.tableCut : MIRROR_CLASS_NAMES.line;
      if (p.role && p.role !== 'table-cut') el.dataset.role = p.role;
      placeAt(el, lineRect(p), offsetY);
      el.setAttribute('aria-hidden', 'true');
      return el;
    }
    case 'image': {
      const el = p.href ? doc.createElement('a') : doc.createElement('div');
      el.className = MIRROR_CLASS_NAMES.image;
      if (p.decorative) {
        el.setAttribute('aria-hidden', 'true');
      } else {
        el.setAttribute('role', 'img');
        if (p.altText) el.setAttribute('aria-label', p.altText);
      }
      applyHrefAttrs(el, p);
      el.dataset.relId = p.relId;
      placeAt(el, { x: p.x, y: p.y, w: p.w, h: p.h }, offsetY);
      applyPrimitiveVisualStyle(el, p.opacity, p.rotationDeg);
      if (p.filter) el.style.filter = p.filter;
      applyDocAttrs(el, p);
      if (p.revision) {
        el.style.outline = `2px solid ${p.revision.kind === 'ins' ? 'rgb(46, 125, 50)' : 'rgb(198, 40, 40)'}`;
      }
      // The mirror is pointer-inert as a whole (page root: pointer-events
      // none), but images keep the painter's contract of a directly
      // hit-testable `.layout-run-image`: re-enabling DOM hit testing lets a
      // click land on the image element itself, then bubble to the canvas
      // host's coordinate-based mousedown routing (imageAtPoint →
      // NodeSelection → resize overlay). Linked images stay pointer-inert so
      // mouse clicks keep the canvas hyperlink routing; their mirror anchors
      // remain keyboard-activatable for a11y either way.
      if (!p.href) el.style.pointerEvents = 'auto';
      return el;
    }
    case 'shape': {
      const el = doc.createElement('div');
      el.className = MIRROR_CLASS_NAMES.shape;
      if (p.decorative) {
        el.setAttribute('aria-hidden', 'true');
      } else {
        el.setAttribute('role', 'img');
      }
      placeAt(el, { x: p.x, y: p.y, w: p.w, h: p.h }, offsetY);
      applyShapeMirrorTransform(el, p.transform);
      applyDocAttrs(el, p);
      return el;
    }
    case 'decoration': {
      const el = doc.createElement('span');
      el.className = MIRROR_CLASS_NAMES.decoration;
      el.dataset.deco = p.deco;
      placeAt(el, { x: p.x, y: p.y, w: p.w, h: p.h }, offsetY);
      applyDocAttrs(el, p);
      return el;
    }
  }
}

function renderPageBorderMirror(
  p: NonNullable<DisplayPage['pageBorders']>[number],
  doc: Document,
  offsetY: number
): HTMLElement {
  const el = doc.createElement('div');
  el.className = MIRROR_CLASS_NAMES.pageBorder;
  placeAt(el, { x: p.x, y: p.y, w: p.w, h: p.h }, offsetY);
  el.style.boxSizing = 'border-box';
  el.style.pointerEvents = 'none';
  el.style.zIndex = p.zOrder === 'back' ? '0' : '20';
  applyPageBorderSideStyle(el, 'Top', p.top);
  applyPageBorderSideStyle(el, 'Right', p.right);
  applyPageBorderSideStyle(el, 'Bottom', p.bottom);
  applyPageBorderSideStyle(el, 'Left', p.left);
  return el;
}

function placeAt(el: HTMLElement, rect: GeoRect, offsetY = 0): void {
  el.style.position = 'absolute';
  el.style.left = `${rect.x}px`;
  el.style.top = `${rect.y + offsetY}px`;
  el.style.width = `${rect.w}px`;
  el.style.height = `${rect.h}px`;
}

function createTextMirrorElement(doc: Document, p: DocAttrs, ctx: MirrorBuildCtx): HTMLElement {
  const noteRef = p.noteRef?.id !== undefined ? p.noteRef : undefined;
  const el = p.href || noteRef ? doc.createElement('a') : doc.createElement('span');
  applyHrefAttrs(el, p);
  if (noteRef && !p.href) {
    // body footnote/endnote reference mark → doc-noteref link to the note's
    // mirror element (W17). Same-document fragment target; id/kind are
    // numeric/enum-derived, so the href cannot carry attacker markup.
    const kind = noteRef.kind ?? 'footnote';
    const target = `oox-${kind}-${noteRef.id}`;
    el.setAttribute('href', `#${target}`);
    el.setAttribute('role', 'doc-noteref');
    // one backlink target per reference, even when the mark split into
    // several primitives (bidi/font subranges)
    const refId = `oox-noteref-${kind}-${noteRef.id}`;
    if (!ctx.noteRefIds.has(refId)) {
      ctx.noteRefIds.add(refId);
      el.id = refId;
    }
  }
  return el;
}

function applyHrefAttrs(el: HTMLElement, p: DocAttrs): void {
  const { href } = p;
  if (!href) return;
  el.setAttribute('href', href);
  el.dataset.href = href;
  const title = p.linkTitle ?? p.tooltip;
  if (title) el.title = title;
  if (p.linkTarget) {
    el.setAttribute('target', p.linkTarget);
  } else if (!href.startsWith('#')) {
    el.setAttribute('target', '_blank');
  }
  if (el.getAttribute('target') === '_blank') {
    el.setAttribute('rel', 'noopener noreferrer');
  }
  if (p.linkHistory !== undefined) el.dataset.linkHistory = String(p.linkHistory);
  if (p.linkDocLocation !== undefined) el.dataset.linkDocLocation = p.linkDocLocation;
}

/**
 * Expose the shaped face's weight/style on a glyph-run mirror span. The canvas
 * paints the (possibly bold/italic) outlines, but the mirror span is the
 * queryable painter-contract surface — browser-shaped runs carry the full font
 * shorthand, so glyph runs must not silently drop the weight. Only fixed enum
 * tokens parsed out of the resolved shorthand reach the CSSOM; the
 * file-derived family string stays out (same policy as the fontSize-only
 * assignment above).
 */
function applyGlyphRunWeightAndStyle(el: HTMLElement, fallbackFont: string | undefined): void {
  if (!fallbackFont) return;
  for (const token of fallbackFont.split(' ')) {
    if (token.endsWith('px')) break; // size reached — the family follows
    if (token === 'italic') el.style.fontStyle = 'italic';
    else if (/^[1-9]00$/.test(token) && token !== '400') el.style.fontWeight = token;
  }
}

function applyRunLanguageAndDirection(
  el: HTMLElement,
  p: { lang?: string; bidiLevel?: number; rtl?: boolean }
): void {
  if (p.lang) el.lang = p.lang;
  if (p.bidiLevel !== undefined) el.dir = p.bidiLevel % 2 === 1 ? 'rtl' : 'ltr';
  else if (p.rtl) el.dir = 'rtl';
}

function applyPageBorderSideStyle(
  el: HTMLElement,
  side: 'Top' | 'Right' | 'Bottom' | 'Left',
  border: { width: number; color: string; style: string } | undefined
): void {
  if (!border) return;
  const style = el.style as CSSStyleDeclaration & Record<string, string>;
  style[`border${side}Width`] = `${border.width}px`;
  style[`border${side}Style`] = border.style;
  style[`border${side}Color`] = border.color;
}

function applyPrimitiveVisualStyle(
  el: HTMLElement,
  opacity?: number,
  rotationDeg?: number,
  horizontalScale?: number
): void {
  if (opacity !== undefined) el.style.opacity = String(opacity);
  const transforms: string[] = [];
  if (rotationDeg) {
    transforms.push(`rotate(${rotationDeg}deg)`);
  }
  if (horizontalScale !== undefined && horizontalScale !== 100) {
    transforms.push(`scaleX(${horizontalScale / 100})`);
  }
  if (transforms.length > 0) {
    el.style.transform = transforms.join(' ');
    el.style.transformOrigin =
      horizontalScale !== undefined && horizontalScale !== 100 ? 'left center' : 'center center';
  }
}

function applyShapeMirrorTransform(
  el: HTMLElement,
  transform:
    | {
        rotation?: number;
        flipH?: boolean;
        flipV?: boolean;
      }
    | undefined
): void {
  if (!transform) return;
  const transforms: string[] = [];
  if (transform.rotation) transforms.push(`rotate(${transform.rotation}deg)`);
  if (transform.flipH) transforms.push('scaleX(-1)');
  if (transform.flipV) transforms.push('scaleY(-1)');
  if (transforms.length === 0) return;
  el.style.transform = transforms.join(' ');
  el.style.transformOrigin = 'center center';
}

function applyTextVisualStyles(
  el: HTMLElement,
  p: {
    allCaps?: boolean;
    smallCaps?: boolean;
    hidden?: boolean;
    textShadow?: 'shadow' | 'emboss' | 'imprint';
    textOutline?: boolean;
    emphasisMark?: 'dot' | 'comma' | 'circle' | 'underDot';
    textEffect?: string;
  }
): void {
  if (p.allCaps) el.style.textTransform = 'uppercase';
  if (p.smallCaps) el.style.fontVariant = 'small-caps';
  if (p.hidden) {
    el.classList.add('docx-hidden');
    el.style.textDecoration = 'underline dotted';
  }
  if (p.textShadow === 'emboss') {
    el.style.textShadow = '1px 1px 1px rgba(255,255,255,0.5), -1px -1px 1px rgba(0,0,0,0.3)';
  } else if (p.textShadow === 'imprint') {
    el.style.textShadow = '-1px -1px 1px rgba(255,255,255,0.5), 1px 1px 1px rgba(0,0,0,0.3)';
  } else if (p.textShadow === 'shadow') {
    el.style.textShadow = '1px 1px 2px rgba(0,0,0,0.3)';
  }
  if (p.textOutline) {
    el.style.webkitTextStroke = '1px currentColor';
    (el.style as CSSStyleDeclaration & { webkitTextFillColor?: string }).webkitTextFillColor =
      'transparent';
  }
  if (p.emphasisMark) {
    const variant =
      p.emphasisMark === 'comma'
        ? 'filled sesame'
        : p.emphasisMark === 'circle'
          ? 'filled circle'
          : 'filled dot';
    const position = p.emphasisMark === 'underDot' ? 'under right' : 'over right';
    el.style.textEmphasis = variant;
    el.style.textEmphasisPosition = position;
    (el.style as CSSStyleDeclaration & { webkitTextEmphasis?: string }).webkitTextEmphasis =
      variant;
    (
      el.style as CSSStyleDeclaration & { webkitTextEmphasisPosition?: string }
    ).webkitTextEmphasisPosition = position;
  }
  if (p.textEffect) {
    el.classList.add('docx-text-effect', `docx-text-effect-${p.textEffect}`);
    el.dataset.effect = p.textEffect;
  }
}

function applyDocAttrs(el: HTMLElement, p: DocAttrs): void {
  if (p.docStart !== undefined) el.dataset.docStart = String(p.docStart);
  if (p.docEnd !== undefined) el.dataset.docEnd = String(p.docEnd);
  if (p.commentIds && p.commentIds.length > 0) {
    el.dataset.commentId = p.commentIds.join(' ');
  }
  // inert field identity: the visible text is the field RESULT; the dataset +
  // description announce what the field is (e.g. `PAGEREF _Toc42 \h`). The
  // instruction is announce-only — the mirror never parses or executes it.
  // Comment/revision descriptions below deliberately override this one.
  if (p.field) {
    const fieldType = p.field.type ?? p.field.category;
    if (fieldType) el.dataset.fieldType = fieldType;
    if (p.field.category) el.dataset.fieldCategory = p.field.category;
    if (p.field.instruction) el.dataset.fieldInstruction = p.field.instruction;
    const description = p.field.instruction?.trim() || fieldType;
    if (description) el.setAttribute('aria-description', description);
  }
  if (p.noteRef?.id !== undefined) {
    el.dataset.noteRefKind = p.noteRef.kind ?? 'footnote';
    el.dataset.noteRefId = String(p.noteRef.id);
  }
  if (p.comment) {
    if (p.comment.status) el.dataset.commentStatus = p.comment.status;
    if (p.comment.authorId) el.dataset.commentAuthorId = p.comment.authorId;
    if (p.comment.paletteIndex !== undefined) {
      el.dataset.commentPaletteIndex = String(p.comment.paletteIndex);
    }
    if (p.comment.color) el.dataset.commentColor = p.comment.color;
    if (p.comment.selected !== undefined) {
      el.dataset.commentSelected = String(p.comment.selected);
    }
    if (p.comment.authorName) el.dataset.commentAuthorName = p.comment.authorName;
    if (p.comment.date) el.dataset.commentDate = p.comment.date;
    if (p.comment.replyCount !== undefined) {
      el.dataset.commentReplyCount = String(p.comment.replyCount);
    }
    // complete-thread announcement: status, author, date, body, then each
    // reply as "author: text". Pure value joins (no hardcoded English), set
    // via setAttribute — the values are file-derived and attacker-controlled.
    const description = [
      p.comment.status,
      p.comment.authorName ?? p.comment.authorId,
      p.comment.date,
      p.comment.text,
      ...(p.comment.replies ?? []).map((reply) =>
        [reply.authorName, reply.text].filter(Boolean).join(': ')
      ),
    ]
      .filter(Boolean)
      .join(', ');
    if (description) el.setAttribute('aria-description', description);
  }
  if (p.revision) {
    el.dataset.changeAuthor = p.revision.author;
    el.dataset.changeDate = p.revision.date;
    el.dataset.revisionId = p.revision.revisionId;
    el.classList.add(p.revision.kind === 'ins' ? 'docx-insertion' : 'docx-deletion');
    // ARIA 1.1 insertion/deletion roles (the ins/del element semantics) so
    // tracked changes are announced; never clobber a stronger role (img)
    if (!el.hasAttribute('role')) {
      el.setAttribute('role', p.revision.kind === 'ins' ? 'insertion' : 'deletion');
    }
    const description = [p.revision.author, p.revision.date].filter(Boolean).join(', ');
    if (description) el.setAttribute('aria-description', description);
  }
  if (p.href) {
    el.dataset.href = p.href;
  }
  applyInlineSdtWidgetAttrs(el, p.inlineSdtWidget);
  applyBlockSdtAttrs(el, p.sdt);
  if (p.sdtPath?.length) el.dataset.sdtPath = p.sdtPath.map((sdt) => sdt.groupId).join(' ');
  if (p.logicalOrder !== undefined) el.dataset.logicalOrder = String(p.logicalOrder);
  if (p.bidiLevel !== undefined) el.dataset.bidiLevel = String(p.bidiLevel);
  applyObjectA11y(el, p);
}

function applyObjectA11y(el: HTMLElement, p: DocAttrs): void {
  if (p.groupId) el.dataset.groupId = p.groupId;
  if (p.decorative || p.hiddenObject) {
    el.setAttribute('aria-hidden', 'true');
    el.removeAttribute('role');
    el.removeAttribute('aria-label');
    el.removeAttribute('aria-description');
    return;
  }
  if (p.ariaLabel) el.setAttribute('aria-label', p.ariaLabel);
  if (p.ariaDescription) el.setAttribute('aria-description', p.ariaDescription);
}

function applyBlockSdtAttrs(el: HTMLElement, sdt: SdtAttrs | undefined): void {
  if (!sdt) return;
  el.classList.add(MIRROR_CLASS_NAMES.blockSdt);
  el.dataset.sdtGroupId = sdt.groupId;
  el.dataset.sdtType = sdt.sdtType;
  if (sdt.depth !== undefined) el.dataset.sdtDepth = String(sdt.depth);
  if (sdt.tag !== undefined) el.dataset.sdtTag = sdt.tag;
  if (sdt.alias !== undefined) el.dataset.sdtAlias = sdt.alias;
  if (sdt.lock !== undefined) el.dataset.sdtLock = sdt.lock;
  if (sdt.checked !== undefined) el.dataset.sdtChecked = String(sdt.checked);
  if (sdt.bound !== undefined) el.dataset.sdtBound = String(sdt.bound);
  if (sdt.repeatingItem !== undefined) el.dataset.sdtRepeatingItem = String(sdt.repeatingItem);
  if (!el.hasAttribute('aria-label') && (sdt.alias || sdt.tag)) {
    el.setAttribute('aria-label', sdt.alias || sdt.tag || '');
  }
  if (sdt.lock) el.setAttribute('aria-readonly', 'true');
}

function applyInlineSdtWidgetAttrs(el: HTMLElement, widget: DocAttrs['inlineSdtWidget']): void {
  if (!widget) return;
  el.classList.add(MIRROR_CLASS_NAMES.inlineSdtWidget);
  const kind = widget.controlKind ?? widget.kind;
  el.dataset.sdtWidget = kind;
  el.dataset.sdtGroupId = widget.groupId;
  el.dataset.sdtPos = String(widget.pos);
  if (widget.tag !== undefined) el.dataset.sdtTag = widget.tag;
  if (widget.alias !== undefined) el.dataset.sdtAlias = widget.alias;
  if (widget.checked !== undefined) {
    el.dataset.sdtChecked = String(widget.checked);
    el.setAttribute('aria-checked', String(widget.checked));
  }
  if (widget.controlId !== undefined) el.dataset.sdtControlId = String(widget.controlId);
  if (widget.value !== undefined) el.dataset.sdtValue = widget.value;
  if (widget.selectedIndex !== undefined) {
    el.dataset.sdtSelectedIndex = String(widget.selectedIndex);
  }
  if (widget.dateFormat !== undefined) el.dataset.sdtDateFormat = widget.dateFormat;
  if (widget.dateLanguage !== undefined) el.dataset.sdtDateLanguage = widget.dateLanguage;
  if (widget.locked !== undefined) el.dataset.sdtLocked = String(widget.locked);
  if (kind === 'checkbox') {
    el.setAttribute('role', 'checkbox');
  } else if (kind === 'dropDownList' || kind === 'comboBox') {
    el.setAttribute('role', 'combobox');
    el.setAttribute('aria-haspopup', 'listbox');
    el.setAttribute('aria-expanded', 'false');
  } else if (kind === 'date') {
    el.setAttribute('role', 'button');
    el.setAttribute('aria-haspopup', 'dialog');
  } else {
    el.setAttribute('role', 'group');
  }
  // The mirror is semantic and pointer-inert. The separately-built interactive
  // overlay owns the only focusable SDT controls.
  el.removeAttribute('tabindex');
  if (widget.alias || widget.tag) el.setAttribute('aria-label', widget.alias || widget.tag || '');
  if (widget.locked) el.setAttribute('aria-disabled', 'true');
}

function applyStructuralRevisionAttrs(el: HTMLElement, revision: StructuralRevision): void {
  switch (revision.scope) {
    case 'pmark':
      el.classList.add(MIRROR_CLASS_NAMES.revisionPmark);
      if (revision.kind === 'ins') el.classList.add(MIRROR_CLASS_NAMES.revisionPmarkIns);
      if (revision.kind === 'del') el.classList.add(MIRROR_CLASS_NAMES.revisionPmarkDel);
      break;
    case 'table':
      el.classList.add(MIRROR_CLASS_NAMES.revisionTable, ooxRevisionKindClass(revision.kind));
      break;
    case 'row':
      el.classList.add(MIRROR_CLASS_NAMES.revisionRow, ooxRevisionKindClass(revision.kind));
      break;
    case 'cell':
      el.classList.add(MIRROR_CLASS_NAMES.revisionCell, ooxRevisionKindClass(revision.kind));
      break;
  }
  el.dataset.revisionId = revision.revisionId;
  el.dataset.revisionAuthor = revision.author;
  if (revision.date) el.dataset.revisionDate = revision.date;
  const description = [revision.author, revision.date].filter(Boolean).join(', ');
  if (description) el.setAttribute('aria-description', description);
}

function ooxRevisionKindClass(kind: StructuralRevision['kind']): string {
  if (kind === 'ins') return MIRROR_CLASS_NAMES.revisionIns;
  if (kind === 'del') return MIRROR_CLASS_NAMES.revisionDel;
  return MIRROR_CLASS_NAMES.revisionMerge;
}

function buildMirrorNoteArea(
  area: NoteRegion,
  page: DisplayPage,
  doc: Document,
  ctx: MirrorBuildCtx
): HTMLElement {
  const el = doc.createElement('section');
  el.className = MIRROR_CLASS_NAMES.notes;
  el.dataset.noteKind = area.kind ?? 'footnote';
  if (area.sectionId) el.dataset.sectionId = area.sectionId;
  if (area.columns !== undefined) el.dataset.columns = String(area.columns);
  el.style.position = 'absolute';
  el.style.left = '0';
  el.style.top = `${area.y ?? 0}px`;
  el.style.width = `${page.width}px`;
  el.style.height = `${area.height ?? 0}px`;

  // W17 backlink metadata by note id (anchor doc range + formatted label)
  const noteMetaById = new Map<number, NoteRegionNote>();
  for (const meta of area.notes ?? []) {
    if (meta.id !== undefined) noteMetaById.set(meta.id, meta);
  }

  appendMirrorPrimitives(el, area.separatorPrimitives ?? [], doc, -(area.y ?? 0), ctx);
  const remaining = [...(area.primitives ?? [])];
  for (const noteId of area.noteIds ?? []) {
    const kind = area.kind ?? 'footnote';
    const groupId = `${kind}-${noteId}`;
    const notePrimitives = remaining.filter((p) => p.groupId === groupId);
    if (notePrimitives.length === 0) continue;
    for (const primitive of notePrimitives) {
      const index = remaining.indexOf(primitive);
      if (index >= 0) remaining.splice(index, 1);
    }
    const note = doc.createElement('aside');
    note.className = MIRROR_CLASS_NAMES.note;
    note.id = `oox-${groupId}`;
    note.dataset.noteId = String(noteId);
    note.setAttribute('role', area.kind === 'endnote' ? 'doc-endnote' : 'doc-footnote');
    const meta = noteMetaById.get(noteId);
    if (meta) {
      if (meta.label !== undefined) note.dataset.noteLabel = meta.label;
      if (meta.anchorDocStart !== undefined) {
        note.dataset.anchorDocStart = String(meta.anchorDocStart);
      }
      if (meta.anchorDocEnd !== undefined) note.dataset.anchorDocEnd = String(meta.anchorDocEnd);
    }
    appendMirrorPrimitives(note, notePrimitives, doc, -(area.y ?? 0), ctx);
    // note → body reference backlink (W17): targets the doc-noteref anchor
    // the body reference mark registered. Labeled by the host-localized
    // MirrorLabels.noteBacklink, falling back to the note's own label/number
    // (data, not hardcoded English). Pointer-inert like the rest of the
    // mirror; navigation UX belongs to the interactive overlay.
    if (meta) {
      const backlink = doc.createElement('a');
      backlink.className = MIRROR_CLASS_NAMES.noteBacklink;
      backlink.setAttribute('role', 'doc-backlink');
      backlink.setAttribute('href', `#oox-noteref-${groupId}`);
      const label = ctx.labels?.noteBacklink ?? meta.label ?? String(noteId);
      backlink.setAttribute('aria-label', label);
      note.appendChild(backlink);
    }
    el.appendChild(note);
  }
  if (remaining.length > 0) appendMirrorPrimitives(el, remaining, doc, -(area.y ?? 0), ctx);
  return el;
}

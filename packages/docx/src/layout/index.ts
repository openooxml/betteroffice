/**
 * Layout — the paged-layout pipeline, one import.
 *
 * A convenience facade over the role-based modules: `flow` (PM doc →
 * FlowBlocks), `measure` (Rust text engine → extents), `pagination` (the
 * pure kernel), `regions` (headers/footers, footnotes, section geometry) and
 * `render` (display list, canvas, geometry queries). Import from a specific
 * module when you know which phase you need; this barrel re-exports the
 * adapter-facing surface of all of them.
 *
 * @experimental The named exports below are the public contract for adapter
 * authors, but the API is still evolving and may change in minor releases
 * until a third-party adapter validates it.
 * @packageDocumentation
 * @public
 */

// PM doc → flow blocks
export {
  toLayoutBlocks,
  resolveListTemplate,
  resetBlockIdCounter,
  convertBorderSpecToLayout,
} from './flow/toLayoutBlocks';
export type { ToLayoutBlocksOptions } from './flow/toLayoutBlocks';

// Table grid + width helpers used by the measurer and paginator.
export {
  resolveTableWidthPx,
  countTableColumns,
  normalizeTableColumnWidths,
  resolveCellGrid,
  resolveTableColumnWidths,
  resolveTableTotalWidthPx,
} from './pagination/tableWidthUtils';
export type { ResolvedGridCell } from './pagination/tableWidthUtils';

// Floating-table classification (demote full-width floats to inline).
export { isBlockLikeFloatingTable, demoteBlockLikeFloatingTables } from './flow/floatingTable';

// Measurement (the Rust source + float pipeline)
export * from './measure';

// Selection-overlay geometry value types. Hit-testing, click→PM position and
// selection rectangles are Rust display-list queries now
// (`layout/render/displayListQueries.ts`, `layout/render/canvasPointer.ts`);
// only these value shapes remain adapter-facing.
export type { SelectionRect, CaretPosition } from './geometry/selectionTypes';

// Footnote layout helpers — full pipeline (page-mapping + content
// conversion via body pipeline) lives in core so any rendering adapter
// (React, Vue, etc.) can share the conversion logic and just supply its
// own platform measureBlocks function.
export {
  collectFootnoteRefs,
  mapFootnotesToPages,
  calculateFootnoteReservedHeights,
  applyFootnotePresentation,
  convertFootnoteToContent,
  buildFootnoteContentMap,
  buildFootnoteRenderItems,
  footnoteReservedHeightsEqual,
  stabilizeFootnoteLayout,
  distributeFootnotesIntoColumns,
  FOOTNOTE_SEPARATOR_HEIGHT,
  FOOTNOTE_COLUMN_GAP_PX,
  MAX_FOOTNOTE_LAYOUT_PASSES,
} from './regions/footnoteLayout';
export type {
  FootnoteRefLocation,
  MeasureBlocksFn,
  ConvertFootnoteOptions,
  StabilizeFootnoteLayoutArgs,
  StabilizeFootnoteLayoutResult,
} from './regions/footnoteLayout';

// Header / footer layout helpers — same pattern as footnote: full pipeline
// (normalization + conversion) lives in core, with adapter-supplied
// `measureBlocks` so the helper stays Canvas-free.
export {
  normalizeHeaderFooterMeasureBlocks,
  resolveHeaderFooterVisualTop,
  calculateHeaderFooterVisualBounds,
  contributesToHeaderFooterFlowHeight,
  convertHeaderFooterToContent,
  computeHfCaretRectsFromDisplayList,
  computeHfSelectionRectsFromDisplayList,
} from './regions/headerFooterLayout';
export type { HeaderFooterMetrics, ConvertHeaderFooterOptions } from './regions/headerFooterLayout';

// Body-margin extension for header/footer band growth. Shared so React + Vue
// pipelines stay in lockstep (issue #705 / #696).
export { extendMarginsForHeaderFooter } from './regions/headerFooterMargins';
export type {
  ExtendMarginsForHeaderFooterInput,
  ExtendMarginsForHeaderFooterResult,
} from './regions/headerFooterMargins';
export { LayoutSelectionGate } from './selectionGate';

// Per-table measurement (recursive over cell content via callback).
export { measureTableBlock, measureTableCellBlockVisualHeight } from './measure/measureTable';

// Section properties → page geometry + header/footer resolution.
export {
  getPageSize,
  getMargins,
  resolveHeaderFooter,
  getColumns,
  columnWidthForSection,
  computePerBlockWidths,
  twipsToPixels,
  DEFAULT_PAGE_WIDTH_PX,
  DEFAULT_PAGE_HEIGHT_PX,
  DEFAULT_BODY_MARGIN_PX,
  DEFAULT_HF_DISTANCE_PX,
} from './regions/sectionGeometry';

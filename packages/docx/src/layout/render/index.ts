/**
 * Display-list rendering surface: the renderer-agnostic contract between
 * layout output and every backend, plus the web canvas replay backend.
 *
 * @experimental The display-list pipeline is under active development
 * (rust-canvas-engine change); shape may evolve while the Rust builder lands.
 */

import type { DisplayList } from './displayList';
import demoDisplayListJson from './__fixtures__/displayList.demo.json';

export type {
  DisplayList,
  DisplayBounds,
  DisplayPage,
  DisplayPrimitive,
  DocAttrs,
  HfRegion,
  TableCellRef,
  TextRunPrimitive,
  GlyphRunPrimitive,
  PositionedGlyph,
  RectPrimitive,
  LinePrimitive,
  ImagePrimitive,
  DecorationPrimitive,
} from './displayList';

export {
  GlyphCache,
  type GlyphCacheOptions,
  type GlyphOutline,
  type GlyphOutlineProvider,
} from './glyphCache';

export { loadGlyphOutlineProvider } from './glyphOutlineWasm';

export {
  buildMirrorPage,
  MIRROR_CLASS_NAMES,
  type BuildMirrorPageOptions,
  type MirrorLabels,
} from './mirrorDom';

export {
  applyInteractiveSdtFocus,
  buildInteractiveOverlayPage,
  type BuildInteractiveOverlayOptions,
  type InteractiveOverlayLabels,
} from './interactiveOverlay';

export {
  drawDisplayPage,
  drawPrimitive,
  presentDisplayPageBackBuffer,
  presentOffscreenPageBackBuffer,
  rasterizeDisplayPageToBackBuffer,
  rasterizeDisplayListPages,
  sizeCanvasForPage,
  type DrawPageOptions,
  type ImageResolver,
  type PageCanvasLike,
  type PageCanvasBuffer,
} from './canvasBackend';

export {
  fontSizePxFromShorthand,
  glyphRunRect,
  lineRect,
  textRunRect,
  type GeoRect,
} from './displayListGeometry';

export {
  buildRustDisplayList,
  buildRustDisplayFrame,
  loadRustDisplayListQueryEngine,
  type DisplayListBuildInputs,
  type DisplayListHeadersFooters,
  type DisplayListHfVariant,
  type DisplayListPictureWatermark,
  type DisplayListTextWatermark,
  type DisplayListWatermark,
  type RustDisplayListEngine,
  type RustDisplayFrameResult,
  type RustDisplayListQueryEngine,
  RustDisplayListSourceError,
  type RustDisplayListSourceErrorStage,
  encodeDisplayListFrameExtras,
} from './rustDisplayList';

export {
  applyFrameDelta,
  applyFrameDeltaOwned,
  decodeFrameDelta,
  FRAME_DELTA_HEADER_BYTES,
  FRAME_DELTA_PAGE_OP_BYTES,
  FRAME_DELTA_VERSION,
  FrameDeltaError,
  type DecodedFrameDelta,
  type FrameDeltaErrorCode,
  type FramePageOperation,
  type RetainedFrame,
  type RetainedFramePage,
} from './frameDelta';

export {
  createDisplayListQueries,
  type DisplayListHitRegion,
  type DisplayListImageGeometry,
  type DisplayListParagraphGeometry,
  type DisplayListQueries,
  type DisplayListQuerySourceState,
  type DisplayListRect,
  type DisplayListRegionHit,
  type DisplayListVerticalMove,
  type DisplayListVisualLine,
  type ResidentDisplayListQueryEngine,
} from './displayListQueries';

export { CANVAS_PAGE_GAP_PX, CANVAS_PAGES_PADDING_PX, canvasPageTops } from './canvasPageMetrics';

export {
  resolveCanvasPoint,
  resolveDisplayPageClientRect,
  type CanvasPointHit,
  type DisplayPageClientRect,
  type DisplayPageHostOptions,
} from './canvasPointer';

export { createCanvasImageResolver } from './canvasImageResolver';

export {
  computeA11yAnnouncements,
  snapshotFromSelectionContext,
  type A11yAnnouncement,
  type A11yCommentThreadDetails,
  type A11ySelectionSnapshot,
} from './a11yAnnouncements';

export {
  computeAnchorPositionsFromDisplayList,
  computeAnchorPositionsFromYrs,
  mergeHfAnchorPositionsFromDisplayList,
  visitAnchorKeys,
  type AnchorEditorView,
  type DisplayListHfAnchorSource,
  type YrsHeaderFooterRegions,
} from './displayListAnchors';

export {
  createYrsSidebarProjection,
  yrsIdToNumericId,
  type YrsSidebarDisplayPoint,
  type YrsSidebarProjection,
} from './yrsSidebarProjection';

export { extractTrackedChangesFromYrs, type TrackedChangesResult } from './yrsTrackedChanges';

export {
  captureInlinePositionEmuFromDisplayList,
  findImagePrimitiveAtPoint,
  findImagePrimitiveByDocPos,
  type DisplayListImageRegion,
  type LocatedImagePrimitive,
} from './displayListImages';

export {
  findDisplayListHyperlinkAtPoint,
  type DisplayListHyperlinkHit,
} from './displayListHyperlinks';

export {
  DISPLAY_LIST_TABLE_INSERT_HIDE_DELAY_MS,
  detectDisplayListTableInsertHover,
  deriveDisplayListSelectedCellRects,
  deriveDisplayListTableFragments,
  type DisplayListTableInsertHoverHit,
  type DisplayListTableInsertHoverInput,
  type DisplayListSelectedCellRect,
  type DisplayListSelectedCellSpec,
  type DisplayListTableFragment,
  type DisplayListTableRegion,
  type TableRowBand,
  type TableKeyResolver,
} from './displayListTables';

/**
 * Demo display list used as the default provider value until the Rust
 * display-list builder plugs in. One page, one paragraph of two runs with a
 * highlight and underline, a table border line, and an image reference.
 */
export const demoDisplayList: DisplayList = demoDisplayListJson as DisplayList;

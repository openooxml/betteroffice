/**
 * @betteroffice/docx (default entry point)
 *
 * Fat barrel that re-exports the parser, serializer, and the most-used types.
 * No React/DOM imports. Adapter authors who only need a specific slice should
 * prefer the smaller subpaths.
 *
 * @example
 * ```ts
 * import { parseDocx, serializeDocx, resolveColor } from '@betteroffice/docx';
 * ```
 * @packageDocumentation
 * @public
 */

// ============================================================================
// VERSION
// ============================================================================

export const VERSION = '0.0.2';

// ============================================================================
// PARSER / SERIALIZER
// ============================================================================

export { parseDocx } from './docx/parser';
export {
  serializeDocument as serializeDocx,
  serializeDocumentBody,
  serializeSectionProperties,
} from './docx/serializer';
export { repackDocx, createDocx, updateMultipleFiles } from './docx/rezip';
export { attemptSelectiveSave } from './docx/selectiveSave';

// ============================================================================
// UTILITIES
// ============================================================================

export {
  twipsToPixels,
  pixelsToTwips,
  formatPx,
  emuToPixels,
  pointsToPixels,
  halfPointsToPixels,
  pixelsToEmu,
  emuToTwips,
  twipsToEmu,
} from './utils/units';

export {
  resolveColor,
  resolveHighlightColor,
  resolveShadingColor,
  parseColorString,
  createThemeColor,
  createRgbColor,
  darkenColor,
  lightenColor,
  blendColors,
  getContrastingColor,
  isBlack,
  isWhite,
  colorsEqual,
  generateThemeTintShadeMatrix,
  getThemeTintShadeHex,
  ensureHexPrefix,
  resolveHighlightToCss,
  type ThemeMatrixCell,
} from './utils/colorResolver';

export { type DocxInput, toArrayBuffer } from './utils/docxInput';

// ============================================================================
// FONT LOADER
// ============================================================================

export {
  loadFont,
  loadFonts,
  loadFontFromBuffer,
  isFontLoaded,
  isLoading as isFontsLoading,
  getLoadedFonts,
  onFontsLoaded,
  canRenderFont,
  preloadCommonFonts,
  setGoogleFontsEnabled,
  isGoogleFontsEnabled,
} from './utils/fontLoader';

export {
  type PrintOptions,
  getDefaultPrintOptions,
  triggerPrint,
  openPrintWindow,
  parsePageRange,
  formatPageRange,
  isPrintSupported,
} from './utils/print';

// ============================================================================
// TYPES
// ============================================================================

export type {
  Document,
  DocxPackage,
  DocumentBody,
  BlockContent,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  TextContent,
  Table,
  TableRow,
  TableCell,
  Image,
  Shape,
  TextBox,
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  Field,
  Theme,
  ThemeColorScheme,
  ThemeFont,
  ThemeFontScheme,
  Style,
  StyleDefinitions,
  TextFormatting,
  ParagraphFormatting,
  SectionProperties,
  HeaderFooter,
  HeaderReference,
  FooterReference,
  Footnote,
  Endnote,
  ListLevel,
  NumberingDefinitions,
  Relationship,
  Comment,
  CommentRangeStart,
  CommentRangeEnd,
  TrackedChangeInfo,
  TrackedRunChange,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
} from './types/document';

// ============================================================================
// EDITOR PLUGIN API (Framework-Agnostic)
// ============================================================================

export type {
  EditorPluginCore,
  PluginPanelProps,
  PanelConfig,
  RenderedDomContext,
  PositionCoordinates,
} from './plugin-api/types';

// ============================================================================
// MANAGER CLASSES (Framework-Agnostic Business Logic)
// ============================================================================

export {
  Subscribable,
  TableSelectionManager,
  ErrorManager,
  TABLE_DATA_ATTRIBUTES,
  findTableFromClick,
  getTableFromDocument,
  updateTableInDocument,
  deleteTableFromDocument,
} from './managers';

export type {
  CellCoordinates,
  TableSelectionSnapshot,
  ErrorSeverity,
  ErrorNotification,
  ErrorManagerSnapshot,
} from './managers';

// ============================================================================
// LAYOUT BRIDGE (Adapter Authoring)
// ============================================================================
//
// Helpers shared by the React + Vue adapters and available to third-party
// adapter authors. The full pipeline (page mapping, content conversion,
// multi-pass convergence) lives in core so every adapter calls the same
// code and stays in lockstep on layout behaviour.

export {
  collectFootnoteRefs,
  mapFootnotesToPages,
  calculateFootnoteReservedHeights,
  buildFootnoteContentMap,
  buildFootnoteRenderItems,
  footnoteReservedHeightsEqual,
  stabilizeFootnoteLayout,
  convertHeaderFooterToContent,
  FOOTNOTE_SEPARATOR_HEIGHT,
  MAX_FOOTNOTE_LAYOUT_PASSES,
} from './layout';

export type {
  FootnoteRefLocation,
  MeasureBlocksFn,
  ConvertFootnoteOptions,
  StabilizeFootnoteLayoutArgs,
  StabilizeFootnoteLayoutResult,
  HeaderFooterMetrics,
  ConvertHeaderFooterOptions,
} from './layout';

export type {
  LayoutBlock,
  Layout,
  BlockExtent,
  Page,
  FootnoteContent,
} from './layout/pagination/types';
export type { ResolvedLine, ResolvedSegment } from './layout/pagination/types';

/**
 * @betteroffice/docx (default entry point)
 *
 * Fat barrel that re-exports the parser, serializer, agent, plugin
 * registry, and the most-used types. No React/DOM imports.
 *
 * **When to import from `.` vs `./headless`:** identical for Node.js
 * use; `.` is the convenient aggregate, `./headless` is its mirror with
 * a slightly different name suffix. Adapter authors who only need a
 * specific slice should prefer the smaller subpaths (`./docx`, `./agent`,
 * `./prosemirror`, `./layout-*`, `./utils`) — they tree-shake better.
 *
 * @example
 * ```ts
 * import { parseDocx, serializeDocx, resolveColor } from '@betteroffice/docx';
 * ```
 * @packageDocumentation
 * @public
 */

/// <reference path="./vueProseMirrorCompat.d.ts" />
/// <reference path="./vueProseMirrorAugment.d.ts" />

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
export { buildPatchedDocumentXml, validatePatchSafety } from './docx/selectiveXmlPatch';

// ============================================================================
// TEMPLATE PROCESSING
// ============================================================================

export {
  processTemplate,
  processTemplateDetailed,
  processTemplateAsBlob,
  getTemplateTags,
  validateTemplate,
  type ProcessTemplateOptions,
  type ProcessTemplateResult,
} from './utils/processTemplate';

// ============================================================================
// DOCUMENT CREATION
// ============================================================================

export {
  createEmptyDocument,
  createDocumentWithText,
  type CreateEmptyDocumentOptions,
} from './utils/createDocument';

// ============================================================================
// AGENT API
// ============================================================================

export { DocumentAgent } from './agent/DocumentAgent';
export { executeCommand, executeCommands } from './agent/executor';
export { getAgentContext, getDocumentSummary, type AgentContextOptions } from './agent/context';
export {
  buildSelectionContext,
  buildExtendedSelectionContext,
  type SelectionContextOptions,
  type ExtendedSelectionContext,
} from './agent/selectionContext';

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

export {
  createPageBreak,
  createColumnBreak,
  createLineBreak,
  createPageBreakRun,
  createPageBreakParagraph,
  insertPageBreak,
  createHorizontalRule,
  insertHorizontalRule,
  isPageBreak,
  isColumnBreak,
  isLineBreak,
  isBreakContent,
  hasPageBreakBefore,
  countPageBreaks,
  findPageBreaks,
  removePageBreak,
  type InsertPosition,
} from './utils/insertOperations';

export { type DocxInput, toArrayBuffer } from './utils/docxInput';

export {
  findStartPosForParaId,
  findParagraphByParaId,
  type ParagraphPositionNode,
} from './utils/paragraphPosition';

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
// VARIABLE DETECTION
// ============================================================================

export {
  detectVariables,
  detectVariablesDetailed,
  detectVariablesInBody,
  detectVariablesInParagraph,
  extractVariablesFromText,
  hasTemplateVariables,
  isValidVariableName,
  sanitizeVariableName,
  formatVariable,
  parseVariable,
  replaceVariables,
  removeVariables,
  documentHasVariables,
  type VariableDetectionResult,
  type VariableOccurrence,
} from './utils/variableDetector';

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
  // Comments + track-changes — also exported from `./headless`, but surfaced
  // here so consumers touching comment threads or revision marks don't have
  // to import from a separate subpath.
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

export type {
  AIAction,
  AIActionRequest,
  AgentResponse,
  AgentContext,
  SelectionContext,
  Range,
  Position,
  ParagraphContext,
  SuggestedAction,
  AgentCommand,
  InsertTextCommand,
  ReplaceTextCommand,
  DeleteTextCommand,
  FormatTextCommand,
  InsertTableCommand,
  InsertImageCommand,
  InsertHyperlinkCommand,
  SetVariableCommand,
  ApplyStyleCommand,
} from './types/agentApi';

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
// CORE PLUGIN SYSTEM
// ============================================================================

export {
  pluginRegistry,
  PluginRegistry,
  registerPlugins,
  docxtemplaterPlugin,
  type CorePlugin,
  type McpToolDefinition,
  type McpToolHandler,
  type McpToolResult,
  type McpSession,
} from './core-plugins';

// ============================================================================
// MANAGER CLASSES (Framework-Agnostic Business Logic)
// ============================================================================

export {
  // Base class
  Subscribable,
  // Manager classes
  AutoSaveManager,
  TableSelectionManager,
  ErrorManager,
  PluginLifecycleManager,
  // AutoSave utilities
  formatLastSaveTime,
  getAutoSaveStatusLabel,
  getAutoSaveStorageSize,
  formatStorageSize,
  isAutoSaveSupported,
  // TableSelection utilities
  TABLE_DATA_ATTRIBUTES,
  findTableFromClick,
  getTableFromDocument,
  updateTableInDocument,
  deleteTableFromDocument,
  // Clipboard utilities
  getSelectionRuns,
  createSelectionFromDOM,
  extractFormattingFromElement,
  rgbToHex,
  // PluginLifecycle utilities
  injectStyles,
  // Coordinators
  LayoutCoordinator,
  EditorCoordinator,
} from './managers';

export type {
  // EditorHandle interface
  EditorHandle,
  // AutoSave types
  AutoSaveStatus,
  AutoSaveManagerOptions,
  SavedDocumentData,
  AutoSaveSnapshot,
  // TableSelection types
  CellCoordinates,
  TableSelectionSnapshot,
  // Error types
  ErrorSeverity,
  ErrorNotification,
  ErrorManagerSnapshot,
  // Plugin types
  PluginLifecycleConfig,
  PluginLifecycleSnapshot,
  // Clipboard types
  ClipboardSelection,
  // LayoutCoordinator types
  SelectionRect,
  CaretPosition,
  ImageSelectionInfo,
  ColumnResizeState,
  LayoutCoordinatorSnapshot,
  // EditorCoordinator types
  EditorLoadingState,
  EditorCoordinatorOptions,
  EditorCoordinatorSnapshot,
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

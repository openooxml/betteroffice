/**
 * Editor utilities (curated public surface).
 *
 * The named exports below are the public API contract. Adding a helper
 * to a source module does not automatically make it public — it must
 * be added to this barrel to be reachable from `@betteroffice/docx/utils`.
 * @packageDocumentation
 * @public
 */

// Unit conversion
export {
  TWIPS_PER_INCH,
  PIXELS_PER_INCH,
  twipsToPixels,
  pixelsToTwips,
  emuToPixels,
  pixelsToEmu,
  emuToTwips,
  twipsToEmu,
  pointsToPixels,
  halfPointsToPixels,
  halfPointsToPoints,
  pointsToHalfPoints,
  eighthsToPixels,
  roundPixels,
  clamp,
  formatPx,
} from './units';

// Color resolution
export {
  resolveColor,
  resolveColorToHex,
  resolveHighlightColor,
  resolveShadingColor,
  isBlack,
  isWhite,
  getContrastingColor,
  parseColorString,
  createThemeColor,
  createRgbColor,
  darkenColor,
  lightenColor,
  blendColors,
  ensureHexPrefix,
  resolveHighlightToCss,
  getThemeTintShadeHex,
  generateThemeTintShadeMatrix,
  colorsEqual,
} from './colorResolver';
export type { ThemeMatrixCell } from './colorResolver';

// Clipboard
export {
  INTERNAL_CLIPBOARD_TYPE,
  CLIPBOARD_TYPES,
  getClipboardImageFiles,
  copyRuns,
  copyParagraphs,
  runsToClipboardContent,
  paragraphsToClipboardContent,
  writeToClipboard,
  readFromClipboard,
  handlePasteEvent,
  parseClipboardHtml,
  isWordHtml,
  isEditorHtml,
  cleanWordHtml,
  htmlToRuns,
  createClipboardHandlers,
} from './clipboard';
export type { ClipboardContent, ParsedClipboardContent, ClipboardOptions } from './clipboard';

// Document creation
export { createEmptyDocument, createDocumentWithText } from './createDocument';
export type { CreateEmptyDocumentOptions } from './createDocument';

// DOCX input handling
export { toArrayBuffer } from './docxInput';
export type { DocxInput } from './docxInput';

// Font loading
export {
  loadFont,
  loadFonts,
  loadFontFromBuffer,
  loadFontFromUrl,
  loadFontDefinitions,
  loadFontWithMapping,
  loadFontsWithMapping,
  preloadCommonFonts,
  loadDocumentFonts,
  isFontLoaded,
  setGoogleFontsEnabled,
  isGoogleFontsEnabled,
  isLoading,
  getLoadedFonts,
  onFontsLoaded,
  onFontError,
  canRenderFont,
  FONT_MAPPING,
  getGoogleFontEquivalent,
  extractFontsFromDocument,
} from './fontLoader';
export type { FontDefinition } from './fontLoader';

// Embedded fonts (de-obfuscation + load + picker discovery)
export { deobfuscateFont, isValidFontKey } from './fontDeobfuscation';
export {
  getEmbeddedFontFaces,
  extractEmbeddedFontFaces,
  loadEmbeddedFonts,
  getEmbeddedFontFamilies,
} from './embeddedFonts';
export type { EmbeddedFontFace } from './embeddedFonts';
export {
  getRenderableDocumentFonts,
  selectRenderableFonts,
  excludeFontsByName,
} from './documentPickerFonts';
export type { RenderableFontOptions } from './documentPickerFonts';

// Formatting → CSS style
export {
  textToStyle,
  paragraphToStyle,
  borderToStyle,
  resolveShadingFill,
  mergeStyles,
  tableCellToStyle,
  sectionToStyle,
} from './formatToStyle';

// Heading collection
export { collectHeadings } from './headingCollector';
export type { HeadingInfo } from './headingCollector';

// Insert operations
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
} from './insertOperations';
export type { InsertPosition } from './insertOperations';

// Keyboard navigation
export {
  isWordCharacter,
  isWhitespace,
  isPunctuation,
  findWordStart,
  findWordEnd,
  findNextWordStart,
  findPreviousWordStart,
  findVisualLineStart,
  findVisualLineEnd,
  getSelectionInfo,
  setSelectionPosition,
  extendSelectionTo,
  moveByWord,
  moveToLineEdge,
  parseNavigationAction,
  handleNavigationKey,
  isNavigationKey,
  expandSelectionToWord,
  getWordAtCursor,
  matchesShortcut,
  describeShortcut,
  getNavigationShortcutDescriptions,
  NAVIGATION_SHORTCUTS,
} from './keyboardNavigation';
export type {
  NavigationDirection,
  NavigationUnit,
  NavigationAction,
  KeyboardShortcut,
} from './keyboardNavigation';

// Template processing
export {
  processTemplate,
  processTemplateDetailed,
  processTemplateAsBlob,
  processTemplateAndDownload,
  getTemplateTags,
  validateTemplate,
  getMissingVariables,
  previewTemplate,
  processTemplateAdvanced,
  createTemplateProcessor,
} from './processTemplate';
export type {
  ProcessTemplateOptions,
  ProcessTemplateResult,
  TemplateError,
} from './processTemplate';

// Selection highlight
export {
  DEFAULT_SELECTION_STYLE,
  HIGH_CONTRAST_SELECTION_STYLE,
  SELECTION_CSS_VARS,
  getSelectionRects,
  mergeAdjacentRects,
  getMergedSelectionRects,
  getHighlightRectStyle,
  generateSelectionCSS,
  hasActiveSelection,
  getSelectedText,
  isSelectionWithin,
  getSelectionBoundingRect,
  highlightTextRange,
  selectRange,
  clearSelection,
  isSelectionBackwards,
  normalizeSelectionDirection,
  injectSelectionStyles,
  removeSelectionStyles,
  areSelectionStylesInjected,
  createSelectionChangeHandler,
} from './selectionHighlight';
export type { HighlightRect, SelectionHighlightConfig, SelectionRange } from './selectionHighlight';

// Paragraph flash helpers
export {
  DEFAULT_PARAGRAPH_FLASH_COLOR,
  DEFAULT_PARAGRAPH_FLASH_DURATION_MS,
  PARAGRAPH_FLASH_CLASS_NAME,
  findParagraphFragmentsByParaId,
  flashParagraphElements,
  flashParagraphFragmentsByParaId,
} from './paragraphFlash';
export type { ParagraphHighlightOptions, ScrollToParaIdOptions } from './paragraphFlash';

// Table split algorithm
export {
  sumColumnWidths,
  redistributeColumnWidths,
  computeSplitLayout,
  buildAnchorMaps,
  computeSplitDialogDefaults,
} from './tableSplitAlgorithm';
export type { CellAnchor, SplitTarget, SplitLayoutResult } from './tableSplitAlgorithm';

// Text selection helpers
export {
  findWordBoundaries,
  getWordAt,
  findWordAt,
  selectWordAtCursor,
  selectWordInTextNode,
  expandSelectionToWordBoundaries,
  selectParagraphAtCursor,
  handleClickForMultiClick,
  createDoubleClickWordSelector,
  createTripleClickParagraphSelector,
} from './textSelection';
export type { WordSelectionResult } from './textSelection';

// Sidebar geometry — shared by both adapters so page-shift + card-gap math stay consistent.
export {
  SIDEBAR_WIDTH,
  SIDEBAR_PAGE_GAP,
  SIDEBAR_DOCUMENT_SHIFT,
  MIN_CARD_GAP,
} from './sidebarConstants';

// File-input reader — shared between every adapter's `<input type=file>`
// → `loadBuffer` glue so filename normalization and the input-reset
// trick can't drift between React, Vue, or any future framework host.
export { readDocxFileFromInput } from './readDocxFile';
export type { ReadDocxFileResult } from './readDocxFile';

// Color-mode (light/dark/auto) resolution — shared so the prefers-color-scheme
// + SSR handling stays identical across React, Vue, and future hosts.
export { prefersColorSchemeDark, resolveIsDark, subscribeSystemDark } from './colorMode';
export type { ColorMode } from './colorMode';

// Comments
export {
  PENDING_COMMENT_ID,
  createCommentIdAllocator,
  type CommentIdAllocator,
} from './commentIdAllocator';
export { createComment } from './commentFactory';

// URL scheme allowlist for hrefs from untrusted input (DOCX rels, pasted HTML).
export { sanitizeHref } from './sanitizeHref';

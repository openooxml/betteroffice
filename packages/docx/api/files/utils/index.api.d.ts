// @betteroffice/docx declaration utils/index.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Editor utilities (curated public surface).
 *
 * The named exports below are the public API contract. Adding a helper
 * to a source module does not automatically make it public — it must
 * be added to this barrel to be reachable from `@betteroffice/docx/utils`.
 * @packageDocumentation
 * @public
 */
export { TWIPS_PER_INCH, PIXELS_PER_INCH, twipsToPixels, pixelsToTwips, emuToPixels, pixelsToEmu, emuToTwips, twipsToEmu, pointsToPixels, halfPointsToPixels, halfPointsToPoints, pointsToHalfPoints, eighthsToPixels, roundPixels, clamp, formatPx, } from './units';
export { resolveColor, resolveColorToHex, resolveHighlightColor, resolveShadingColor, isBlack, isWhite, getContrastingColor, parseColorString, createThemeColor, createRgbColor, darkenColor, lightenColor, blendColors, ensureHexPrefix, resolveHighlightToCss, getThemeTintShadeHex, generateThemeTintShadeMatrix, colorsEqual, } from './colorResolver';
export type { ThemeMatrixCell } from './colorResolver';
export { toArrayBuffer } from './docxInput';
export type { DocxInput } from './docxInput';
export { loadFont, loadFonts, loadFontFromBuffer, loadFontFromUrl, loadFontDefinitions, loadFontWithMapping, loadFontsWithMapping, preloadCommonFonts, loadDocumentFonts, isFontLoaded, setGoogleFontsEnabled, isGoogleFontsEnabled, isLoading, getLoadedFonts, onFontsLoaded, onFontError, canRenderFont, FONT_MAPPING, getGoogleFontEquivalent, extractFontsFromDocument, createFontLoadScope, registerDocumentFaces, } from './fontLoader';
export type { BufferFaceInput, FontDefinition, FontLoadScope } from './fontLoader';
export { deobfuscateFont, isValidFontKey } from './fontDeobfuscation';
export { getEmbeddedFontFaces, extractEmbeddedFontFaces, loadEmbeddedFonts, loadEmbeddedFontFamilies, getEmbeddedFontFamilies, } from './embeddedFonts';
export type { EmbeddedFontFace } from './embeddedFonts';
export { getRenderableDocumentFonts, selectRenderableFonts, excludeFontsByName, } from './documentPickerFonts';
export type { RenderableFontOptions } from './documentPickerFonts';
export { collectHeadings } from './headingCollector';
export type { HeadingInfo } from './headingCollector';
export { DEFAULT_PARAGRAPH_FLASH_COLOR, DEFAULT_PARAGRAPH_FLASH_DURATION_MS, PARAGRAPH_FLASH_CLASS_NAME, findParagraphFragmentsByParaId, flashParagraphElements, flashParagraphFragmentsByParaId, } from './paragraphFlash';
export type { ParagraphHighlightOptions, ScrollToParaIdOptions } from './paragraphFlash';
export { sumColumnWidths, redistributeColumnWidths, computeSplitLayout, buildAnchorMaps, computeSplitDialogDefaults, } from './tableSplitAlgorithm';
export type { CellAnchor, SplitTarget, SplitLayoutResult } from './tableSplitAlgorithm';
export { findWordBoundaries, getWordAt, findWordAt, selectWordAtCursor, selectWordInTextNode, expandSelectionToWordBoundaries, selectParagraphAtCursor, handleClickForMultiClick, createDoubleClickWordSelector, createTripleClickParagraphSelector, } from './textSelection';
export type { WordSelectionResult } from './textSelection';
export { SIDEBAR_WIDTH, SIDEBAR_PAGE_GAP, SIDEBAR_DOCUMENT_SHIFT, MIN_CARD_GAP, } from './sidebarConstants';
export { readDocxFileFromInput } from './readDocxFile';
export type { ReadDocxFileResult } from './readDocxFile';
export { prefersColorSchemeDark, resolveIsDark, subscribeSystemDark } from './colorMode';
export type { ColorMode } from './colorMode';
export { PENDING_COMMENT_ID, createCommentIdAllocator, type CommentIdAllocator, } from './commentIdAllocator';
export { createComment } from './commentFactory';
export { sanitizeHref } from './sanitizeHref';

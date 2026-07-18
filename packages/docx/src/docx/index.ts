/**
 * DOCX I/O
 *
 * Parsing DOCX archives into the `Document` model and re-zipping a
 * model back into a DOCX file. Use `./docx/serializer` for the lower-level
 * Document → XML transforms.
 *
 * The named exports below are the public API contract. Adding a parser
 * helper to a source module does not automatically make it public — it
 * must be added to this barrel to be reachable from
 * `@betteroffice/docx/docx`.
 * @packageDocumentation
 * @public
 */

// Top-level archive I/O
export { parseDocx } from './parser';
export { parseFontTable } from './fontTableParser';
export { repackDocx, createDocx, updateMultipleFiles } from './rezip';
export { attemptSelectiveSave } from './selectiveSave';
export { buildPatchedDocumentXml, validatePatchSafety } from './selectiveXmlPatch';

// Image / drawing
export {
  emuToPixels,
  pixelsToEmu,
  parseDrawing,
  parseImage,
  isInlineImage,
  isFloatingImage,
  isBehindText,
  isInFrontOfText,
  getImageWidthPx,
  getImageHeightPx,
  getWrapDistancesPx,
  isDecorativeImage,
} from './imageParser';
export {
  calculateResizedImageDimensions,
  deriveLayoutChoice,
  IMAGE_LAYOUT_OPTIONS,
  isImageLayoutOptionEnabled,
  resolveImageLayoutAttrs,
  toolbarValueToLayoutTarget,
} from './imageLayout';
export type {
  AnchorWrapType,
  ImageLayoutAttrs,
  ImageLayoutCurrentAttrs,
  ImageLayoutIconHint,
  ImageLayoutOptionDef,
  ImageLayoutPosition,
  ImageLayoutTarget,
  ImageResizeHandle,
  SetImageWrapTypeOptions,
} from './imageLayout';

// Footnotes / endnotes
export {
  parseFootnotes,
  parseEndnotes,
  parseFootnoteProperties,
  parseEndnoteProperties,
  getFootnoteText,
  getEndnoteText,
  isSeparatorFootnote,
  isSeparatorEndnote,
} from './footnoteParser';
export type { FootnoteMap, EndnoteMap } from './footnoteParser';

// Fields
export {
  KNOWN_FIELD_TYPES,
  parseFieldType,
  isKnownFieldType,
  parseFieldInstruction,
  getFormatSwitch,
  hasMergeFormat,
  parseSimpleField,
  createComplexFieldContext,
  isPageNumberField,
  isTotalPagesField,
  isDateTimeField,
  isDocPropertyField,
  isReferenceField,
  isMergeField,
  isTocField,
  getFieldDisplayValue,
} from './fieldParser';
export type {
  ParsedFieldInstruction,
  FieldSwitch,
  ComplexFieldState,
  ComplexFieldContext,
} from './fieldParser';

// Hyperlinks
export {
  parseHyperlink,
  getHyperlinkText,
  isExternalLink,
  isInternalLink,
  getHyperlinkUrl,
  hasContent,
  getHyperlinkRuns,
  resolveHyperlinkUrl,
} from './hyperlinkParser';

// Numbering / lists
export {
  parseNumbering,
  createNumberingMap,
  getCachedNumberingMap,
  computeListRendering,
  formatNumber,
  renderListMarker,
  getBulletCharacter,
  isBulletLevel,
} from './numberingParser';
export type { NumberingMap } from './numberingParser';

// Shapes
export {
  parseShape,
  parseShapeFromDrawing,
  isShapeDrawing,
  isLineShape,
  isTextBoxShape,
  hasTextContent,
  getShapeWidthPx,
  isFloatingShape,
  hasFill,
  hasOutline,
  getOutlineWidthPx,
  resolveFillColor,
  resolveOutlineColor,
  getShapeHeightPx,
  getShapeDimensionsPx,
} from './shapeParser';

// Tab stops
export {
  DEFAULT_TAB_INTERVAL_TWIPS,
  DEFAULT_TAB_ALIGNMENT,
  DEFAULT_TAB_LEADER,
  parseTabStop,
  parseTabStops,
  parseTabStopsFromParagraphProperties,
  mergeTabStops,
  getNextTabStop,
  calculateTabWidth,
  calculateTabWidthWithAlignment,
  getLeaderCharacter,
  hasVisibleLeader,
} from './tabParser';

// Tables
export {
  parseTableMeasurement,
  parseBorderSpec,
  parseTableBorders,
  parseCellMargins,
  parseShading,
  parseTableLook,
  parseFloatingTableProperties,
  parseTableProperties,
  getTableColumnCount,
  getTableRowCount,
  isCellMergeContinuation,
  isCellMergeStart,
  hasHeaderRow,
} from './tableParser';

// Text boxes
export {
  extractTextBoxContentElements,
  parseTextBoxContent,
  isTextBoxDrawing,
  isShapeTextBox,
  parseTextBox,
  getTextBoxContentElement,
  parseTextBoxFromShape,
  getTextBoxWidthPx,
  getTextBoxHeightPx,
  getTextBoxDimensionsPx,
  getTextBoxMarginsPx,
  isFloatingTextBox,
  hasTextBoxFill,
  hasTextBoxOutline,
  hasTextBoxContent,
  getTextBoxText,
  resolveTextBoxFillColor,
  resolveTextBoxOutlineColor,
  getTextBoxOutlineWidthPx,
} from './textBoxParser';
export type { ParagraphParserFn, TableParserFn } from './textBoxParser';

// Reply-range marker injection — pre-serialization step that
// synthesizes commentRange markers for reply comments. Pure data
// transform; both adapters call it before saving.
export { injectReplyRangeMarkers, injectTCReplyRangeMarkers } from './injectReplyRangeMarkers';
export { getDocumentWatermark, setDocumentWatermark } from './watermarkApi';

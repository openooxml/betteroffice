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
export { repackDocx, createDocx, updateMultipleFiles } from './rezip';
export { attemptSelectiveSave } from './selectiveSave';

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

// Reply-range marker injection — pre-serialization step that
// synthesizes commentRange markers for reply comments. Pure data
// transform; both adapters call it before saving.
export { injectReplyRangeMarkers, injectTCReplyRangeMarkers } from './injectReplyRangeMarkers';
export { getDocumentWatermark, setDocumentWatermark } from './watermarkApi';

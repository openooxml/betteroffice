import type { ProjectionAllowance } from './projectionParity.allowlist';

export const fields: readonly ProjectionAllowance[] = [
  { kind: 'field', path: 'blocks[].content[]', reason: 'Field lowering is deferred.' },
  { kind: 'field', path: 'blocks[].content[].content[]', reason: 'The source field code can retain its instruction or delimiter run.' },
  { kind: 'hyperlink', path: 'blocks[].content[]', reason: 'Hyperlink marks are deferred.' },
  { kind: 'hyperlink', path: 'blocks[].content[].href', reason: 'A retained local hyperlink can still carry its source href.' },
  { kind: 'pageBreak', path: 'blocks[].content[].content[]', reason: 'Page break lowering is deferred.' },
  { kind: 'noteRef', path: 'blocks[].content[].content[].customMarkFollows', reason: 'A retained note reference can carry the source custom-mark flag.' },
  { kind: 'image', path: 'blocks[].content[].content[].drawing', reason: 'Retained pictures carry their source drawing properties.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.filename', reason: 'Retained pictures carry the source filename.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.mimeType', reason: 'Retained pictures carry the source MIME type.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.id', reason: 'Retained pictures carry the source drawing ID.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.name', reason: 'Retained pictures carry the source drawing name.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.rotationBounds', reason: 'Retained pictures carry the source rotation bounds.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.outline', reason: 'Retained pictures carry the source outline.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.position.simplePos', reason: 'Retained pictures carry the source simple position.' },
  { kind: 'image', path: 'blocks[].content[].content[].image.position.horizontal.offset', reason: 'Retained pictures carry the source horizontal offset.' },
  { kind: 'opaqueDrawing', path: 'blocks[].content[].content[].preview.filename', reason: 'Retained opaque drawings carry the source preview filename.' },
  { kind: 'opaqueDrawing', path: 'blocks[].content[].content[].preview.mimeType', reason: 'Retained opaque drawings carry the source preview MIME type.' },
  { kind: 'opaqueDrawing', path: 'blocks[].content[].content[].drawing', reason: 'Retained opaque drawings carry their source drawing properties.' },
  { kind: 'shape', path: 'blocks[].content[].content[].drawing', reason: 'Retained shapes carry their source drawing properties.' },
];

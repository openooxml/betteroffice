/**
 * Read-only structured export of PPTX content: slides in deck order, each with its shape tree,
 * text stories, tables and placeholders for content the export cannot represent, every record
 * with the location it was read from, and diagnostics for everything omitted. Plain JSON,
 * produced by one Rust walker for sessions and bytes.
 *
 * Reading order is the current shape tree, depth first: slides in deck order, shapes in tree
 * order, a group's descendants at the group's position, table cells row by row. It is the
 * authored order, not a reading order inferred from geometry.
 *
 * `range` anchors are batch text targets in the story offsets of `readContent()`, so a session
 * export's `range` anchor can be a step's `target` at the version it was read at; `notes` and
 * `comment` anchors index the plain text they carry; `sourcePart` anchors address retained source
 * XML by element-child ordinals. Session anchors belong to the version they were read at; ids are
 * deterministic export-tree paths.
 */

import type { OperationRefusal } from '../../../shared/host-contracts/edits';
import type { PptxTextTarget } from './edits';
import type {
  ExportCompletion,
  ExportDiagnostic,
  MarkdownAnchor,
} from '../../../shared/host-contracts/exports';

export type {
  ExportCompletion,
  ExportDiagnostic,
  ExportSeverity,
  MarkdownAnchor,
} from '../../../shared/host-contracts/exports';

/** `session` anchors resolve against the returned version; `snapshot` ones against the content. */
export type PptxAnchorScope = 'session' | 'snapshot';

/** A half-open range of UTF-16 offsets. */
export interface PptxTextSpan {
  start: number;
  end: number;
}

export type PptxAnchor =
  | { kind: 'slide'; slideId: string }
  | { kind: 'shape'; slideId: string; shapeId: string }
  | Extract<PptxTextTarget, { kind: 'range' }>
  | { kind: 'notes'; slideId: string; range: PptxTextSpan }
  | { kind: 'comment'; slideId: string; commentId: string; range: PptxTextSpan }
  | { kind: 'sourcePart'; part: string; partSha256: string; path: number[] };

/**
 * The retained source XML a record was seeded from: `sldId` is the owning slide's `p:sldId/@id`,
 * `sourceId` a shape's `cNvPr/@id`.
 */
export interface PptxSourceProvenance {
  part: string;
  partSha256: string;
  path: number[];
  sldId: number | null;
  sourceId: number | null;
}

export interface PptxStructuredContent extends ExportCompletion {
  schemaVersion: 1;
  anchorScope: PptxAnchorScope;
  readingOrder: 'shapeTree';
  included: {
    hiddenSlides: boolean;
    hiddenShapes: boolean;
    notes: boolean;
    comments: boolean;
    formatting: boolean;
  };
  slides: PptxExportSlide[];
  diagnostics: PptxExportDiagnostic[];
}

export interface PptxExportSlide {
  id: string;
  /** Zero-based position in the current deck, hidden slides counted. */
  index: number;
  anchor: PptxAnchor;
  name: string | null;
  /**
   * `null` when the deck does not record it: such slides are exported with a
   * `visibility-unknown` diagnostic.
   */
  hidden: boolean | null;
  provenance: PptxSourceProvenance | null;
  shapes: PptxExportShape[];
  notes: PptxExportNotes | null;
  comments: PptxExportComment[];
}

/** `unknown` is a shape-tree element only the source holds, anchored by its source location. */
export type PptxExportShapeKind = 'shape' | 'picture' | 'graphicFrame' | 'group' | 'unknown';

export interface PptxExportShape {
  id: string;
  anchor: PptxAnchor;
  kind: PptxExportShapeKind;
  name: string;
  /** Alternative-text title and description. */
  title: string | null;
  description: string | null;
  /** Hidden itself or through a hidden group. */
  hidden: boolean;
  placeholder: { type: string | null; index: number | null } | null;
  provenance: PptxSourceProvenance | null;
  stories: PptxExportStory[];
  table: PptxExportTable | null;
  object: PptxExportObject | null;
  children: PptxExportShape[];
}

export interface PptxExportStory {
  id: string;
  /** Spans the whole story. */
  anchor: PptxAnchor;
  paragraphs: PptxExportParagraph[];
}

export interface PptxExportParagraph {
  id: string;
  anchor: PptxAnchor;
  paragraphId: string;
  level: number;
  /** The authored alignment, else the inherited one. */
  alignment: string | null;
  /** The paragraph's own bullet, as the session stores it. */
  bulletJson: string | null;
  /** The marker after inheritance and numbering; `null` for a paragraph that is no list item. */
  list: PptxExportList | null;
  runs: PptxExportRun[];
}

/**
 * `character` is the authored `a:buChar` and `font` its typeface, a symbol font when the character
 * names a glyph slot. `marker` is the formatted number, `null` for a scheme that cannot be
 * formatted.
 */
export type PptxExportList =
  | { kind: 'bullet'; character: string; font: string | null }
  | { kind: 'number'; scheme: string; startAt: number; value: number; marker: string | null };

export type PptxExportMark =
  | { kind: 'bold' | 'italic' | 'superscript' | 'subscript' | 'smallCaps' | 'allCaps' }
  | { kind: 'underline'; style: string };

/** Line breaks are one `\n` unit, fields cover their cached result; unsupported runs are zero-width. */
export type PptxExportRun = {
  anchor: PptxAnchor;
  /** `null` when the export excludes formatting. */
  marks: PptxExportMark[] | null;
  /** A click hyperlink: an external URL, or the part an internal link jumps to. */
  link: { href: string; external: boolean } | null;
} & (
  | { kind: 'text'; text: string }
  | { kind: 'lineBreak' }
  /** A field's cached result; fields are never evaluated. */
  | { kind: 'field'; fieldType: string | null; text: string }
  | { kind: 'unsupported'; element: string }
);

/**
 * A table on its grid. Every source cell appears once; `merged` marks a continuation covered by
 * the cell at `mergeOrigin`, whose content PowerPoint does not show.
 */
export interface PptxExportTable {
  columns: number;
  rows: Array<{ cells: PptxExportTableCell[] }>;
}

export interface PptxExportTableCell {
  id: string;
  anchor: PptxAnchor;
  row: number;
  column: number;
  gridSpan: number;
  rowSpan: number;
  merged: boolean;
  mergeOrigin: { row: number; column: number } | null;
  story: PptxExportStory | null;
}

export type PptxExportObjectKind =
  | 'picture'
  | 'video'
  | 'audio'
  | 'chart'
  | 'smartArt'
  | 'embeddedObject'
  | 'unknown';

/** A placeholder for content the export does not represent; its data is never included. */
export interface PptxExportObject {
  id: string;
  kind: PptxExportObjectKind;
  /** The source element's qualified name, such as `p:graphicFrame`. */
  element: string;
  uri: string | null;
  relationshipIds: string[];
  /** The package parts the object references. */
  parts: string[];
  /** Targets of its external relationships, such as a linked video. */
  externalTargets: string[];
}

/** A slide's speaker notes, as plain text. */
export interface PptxExportNotes {
  id: string;
  anchor: PptxAnchor;
  text: string;
  provenance: PptxSourceProvenance | null;
}

export interface PptxExportComment {
  id: string;
  anchor: PptxAnchor;
  commentId: string;
  author: string | null;
  date: string | null;
  parentId: string | null;
  resolved: boolean;
  text: string;
}

export type PptxExportDiagnosticCode =
  | 'unsupported-content'
  | 'image-data-omitted'
  | 'unsupported-numbering'
  | 'field-cached-result'
  | 'provenance-unavailable'
  | 'visibility-unknown'
  | 'hidden-content-excluded'
  | 'stories-omitted'
  | 'formatting-omitted'
  | 'inherited-content-omitted'
  | 'notes-structure-omitted'
  | 'merge-continuation-content-omitted'
  | 'truncated'
  | 'markdown-lossy';

export type PptxExportDiagnostic = ExportDiagnostic<PptxExportDiagnosticCode, PptxAnchor>;

/**
 * What to export. Hidden slides and shapes, notes and comments are excluded unless requested, and
 * formatting is included unless turned off. `maxBlocks` defaults to 10,000 (at most 1,000,000;
 * every slide, shape, paragraph, notes story and comment counts) and `maxBytes` to 8,388,608
 * (1,024 to 67,108,864, measured on the compact JSON).
 */
export interface PptxExportOptions {
  includeHiddenSlides?: boolean;
  includeHiddenShapes?: boolean;
  includeNotes?: boolean;
  includeComments?: boolean;
  includeFormatting?: boolean;
  maxBlocks?: number;
  maxBytes?: number;
}

/** `maxBytes` bounds the Markdown text; the default and range are the export's. */
export interface PptxMarkdownOptions {
  maxBytes?: number;
}

/**
 * Markdown rendered from structured content. Each `<!-- pptx-export:N -->` marker precedes the
 * record `anchors` maps it to. Markdown keeps no slide layout, geometry or typography.
 */
export interface PptxMarkdownContent extends ExportCompletion {
  markdown: string;
  anchors: MarkdownAnchor<PptxAnchor>[];
  diagnostics: PptxExportDiagnostic[];
}

/**
 * Options out of range or above a hard maximum, or structured content handed to the renderer that
 * contradicts itself. Unsupported deck content is diagnosed, never refused.
 */
export type PptxExportFailureCode = 'invalid-options' | 'limit-exceeded' | 'invalid-content';

/** Why an export was refused; `target` is the anchor the refusal concerns, if any. */
export interface PptxExportFailure {
  code: PptxExportFailureCode;
  target: PptxAnchor | null;
  message: string;
}

export type PptxExportRefusal = OperationRefusal<PptxExportFailure>;

/** A session export: the content with the version it was read at, or a refusal. */
export type PptxExportResult<T> = { ok: true; version: string; content: T } | PptxExportRefusal;

/** A headless export or rendering the engine refused as data. */
export class PptxExportError extends Error {
  readonly failure: PptxExportFailure;

  constructor(failure: PptxExportFailure) {
    super(failure.message);
    this.name = 'PptxExportError';
    this.failure = failure;
  }
}

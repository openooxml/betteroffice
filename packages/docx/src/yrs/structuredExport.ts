/**
 * Read-only structured export of DOCX content: ordered stories of typed blocks and inlines, each
 * with the location it was read from, and diagnostics for everything omitted or unrepresentable.
 * Plain JSON, produced by one Rust walker for sessions, bytes and the native facade.
 *
 * Text, tab and atom inlines carry `range` anchors into the projected text of their view (the
 * batch projection: tabs are `\t`, every other atom one U+FFFC). In the markup view inserted and
 * unchanged text use accepted-view ranges and deleted text original-view ranges. Children of an
 * atom carry the atom's or control's anchor. Ids are deterministic export-tree paths.
 */

import type { OperationRefusal } from '../../../../shared/host-contracts/edits';
import type {
  ExportCompletion,
  ExportDiagnostic,
  MarkdownAnchor,
} from '../../../../shared/host-contracts/exports';
import type {
  DocxAnchor,
  DocxControlMetadata,
  DocxHeadingInfo,
  DocxStorySelection,
} from './readTypes';

export type DocxRevisionView = 'accepted' | 'original' | 'markup';

/** `session` anchors resolve against the returned version; `snapshot` ones against the content. */
export type DocxAnchorScope = 'session' | 'snapshot';

export type DocxStoryKind = 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'comment';

export interface DocxStructuredContent extends ExportCompletion {
  schemaVersion: 1;
  revisionView: DocxRevisionView;
  anchorScope: DocxAnchorScope;
  includedStories: DocxStorySelection[];
  includeFormatting: boolean;
  stories: DocxExportStory[];
  diagnostics: DocxExportDiagnostic[];
}

/** One story. Table cells and block controls nest inside the blocks that own them. */
export interface DocxExportStory {
  story: string;
  kind: DocxStoryKind;
  part: string | null;
  noteId: string | null;
  comment: DocxCommentMetadata | null;
  /** The sections referencing a header or footer part, inheritance applied. */
  uses: Array<{ sectionIndex: number; variant: 'default' | 'first' | 'even' }>;
  blocks: DocxExportBlock[];
}

export interface DocxCommentMetadata {
  id: string;
  author: string | null;
  date: string | null;
  parentId: string | null;
  resolved: boolean;
  /** The ranges the comment annotates. */
  anchors: DocxAnchor[];
}

export interface DocxParagraphData {
  styleId: string | null;
  inlines: DocxExportInline[];
}

/** `format` is the OOXML numbering format; `marker` the rendered marker, `null` if unresolved. */
export interface DocxListInfo {
  numId: string;
  abstractNumId: string | null;
  level: number;
  format: string;
  marker: string | null;
  suffix: 'tab' | 'space' | 'nothing';
  markerHidden: boolean;
}

export type DocxExportBlock = {
  id: string;
  anchor: DocxAnchor;
} & (
  | { kind: 'paragraph'; paragraph: DocxParagraphData }
  | { kind: 'heading'; paragraph: DocxParagraphData; heading: DocxHeadingInfo }
  /** A numbered heading is a list item with a heading. */
  | {
      kind: 'listItem';
      paragraph: DocxParagraphData;
      list: DocxListInfo;
      heading: DocxHeadingInfo | null;
    }
  | { kind: 'table'; table: DocxExportTable }
  | {
      kind: 'contentControl';
      control: DocxControlMetadata;
      story: string | null;
      blocks: DocxExportBlock[];
    }
  /** The break that starts section `sectionIndex`, typed by that section's start. */
  | {
      kind: 'sectionBreak';
      sectionIndex: number;
      breakType: 'continuous' | 'nextPage' | 'nextColumn' | 'oddPage' | 'evenPage';
    }
  /** A page or column break between blocks. */
  | { kind: 'break'; breakType: 'page' | 'column' }
  | { kind: 'unsupported'; element: string }
);

/**
 * A table on its zero-based grid. Every covered grid position of a row appears once, as a cell
 * owning content or as a vertical-merge continuation (`rowSpan: 0`) naming its origin.
 */
export interface DocxExportTable {
  gridColumns: number;
  rows: Array<{
    header: boolean;
    gridBefore: number;
    gridAfter: number;
    cells: DocxExportCell[];
  }>;
}

export interface DocxExportCell {
  anchor: DocxAnchor;
  story: string | null;
  column: number;
  gridSpan: number;
  rowSpan: number;
  verticalMerge: 'none' | 'restart' | 'continue';
  mergeOrigin: { row: number; column: number } | null;
  blocks: DocxExportBlock[];
}

export type DocxExportInline = {
  id: string;
  anchor: DocxAnchor;
  /** `null` when the export excludes formatting. */
  marks: DocxFormattingMark[] | null;
  link: { href: string; title: string | null } | null;
  /** Revision attribution; only the markup view carries any. */
  revisions: DocxInlineRevision[];
} & (
  | { kind: 'text'; text: string }
  | { kind: 'tab' }
  | { kind: 'break'; breakType: 'line' | 'page' | 'column' }
  | { kind: 'noteReference'; noteKind: 'footnote' | 'endnote'; noteId: string; story: string | null }
  | { kind: 'commentReference'; commentId: string; story: string | null }
  /** A field with its cached result; never evaluated. */
  | {
      kind: 'field';
      fieldType: string;
      instruction: string;
      cachedResult: DocxCachedResult;
      dirty: boolean;
      locked: boolean;
    }
  | {
      kind: 'image';
      altText: string | null;
      relationshipId: string | null;
      part: string | null;
      externalTarget: string | null;
    }
  | { kind: 'contentControl'; control: DocxControlMetadata; inlines: DocxExportInline[] }
  | { kind: 'unsupported'; element: string; altText: string | null }
);

export type DocxCachedResult =
  | { kind: 'missing' }
  | { kind: 'inline'; inlines: DocxExportInline[] }
  | { kind: 'blocks'; blocks: DocxExportBlock[] };

export type DocxFormattingMark =
  | { kind: 'bold' | 'italic' | 'strike' | 'subscript' | 'superscript' | 'hidden' }
  | { kind: 'underline'; style: string };

export interface DocxInlineRevision {
  kind: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo';
  id: string | null;
  author: string | null;
  date: string | null;
}

export type DocxExportDiagnosticCode =
  | 'unsupported-content'
  | 'unsupported-revision'
  | 'unsupported-numbering'
  | 'unresolved-style'
  | 'unresolved-reference'
  | 'provenance-unavailable'
  | 'ambiguous-identity'
  | 'image-data-omitted'
  | 'field-cached-result'
  | 'missing-field-result'
  | 'formatting-omitted'
  | 'stories-omitted'
  | 'revision-content-excluded'
  | 'merge-continuation-content-omitted'
  | 'legacy-control-value'
  | 'parse-warning'
  | 'truncated'
  | 'markdown-lossy';

export type DocxExportDiagnostic = ExportDiagnostic<DocxExportDiagnosticCode, DocxAnchor>;

/**
 * What to export. `revisionView` is required; `stories` defaults to `['body']`,
 * `includeFormatting` to `true`, `maxBlocks` to 10,000 (at most 1,000,000; every nested block
 * counts) and `maxBytes` to 8,388,608 (1,024 to 67,108,864, measured on the compact JSON).
 */
export interface DocxExportOptions {
  revisionView: DocxRevisionView;
  stories?: readonly DocxStorySelection[];
  includeFormatting?: boolean;
  maxBlocks?: number;
  maxBytes?: number;
}

/** `maxBytes` bounds the Markdown text; the default and range are the export's. */
export interface DocxMarkdownOptions {
  maxBytes?: number;
}

/**
 * Markdown rendered from structured content. Each `<!-- docx-export:N -->` marker precedes the
 * block `anchors` maps it to. Markdown keeps no Word pagination, typography or layout.
 */
export interface DocxMarkdownContent extends ExportCompletion {
  markdown: string;
  anchors: MarkdownAnchor<DocxAnchor>[];
  diagnostics: DocxExportDiagnostic[];
}

/**
 * Options out of range or above a hard maximum, or a session holding no document content.
 * Unsupported document content is diagnosed, never refused.
 */
export type DocxExportFailureCode = 'invalid-options' | 'limit-exceeded' | 'unsupported';

/** Why an export was refused; `target` is the anchor the refusal concerns, if any. */
export interface DocxExportFailure {
  code: DocxExportFailureCode;
  target: DocxAnchor | null;
  message: string;
}

export type DocxExportRefusal = OperationRefusal<DocxExportFailure>;

/** A session export: the content with the version it was read at, or a refusal. */
export type DocxExportResult<T> = { ok: true; version: string; content: T } | DocxExportRefusal;

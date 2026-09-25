/**
 * Read types shared by the DOCX read APIs: where read content lives, which stories a read
 * covers, heading classification and content-control metadata. Plain JSON, produced by Rust.
 */

import type { DocxTextRange } from './edits';

/**
 * Where a piece of read content lives. Paragraph, range, table and control anchors resolve
 * against the version or snapshot they were read at; table indices and control ids are not stable
 * across saves. A `sourcePart` anchor addresses retained source XML (the part, its SHA-256 and
 * zero-based element-child ordinals from the part's root); it is provenance, not an edit target.
 * A paragraph anchor with an empty `paraId` marks content with no location of its own.
 */
export type DocxAnchor =
  | { kind: 'paragraph'; story: string; paraId: string }
  | ({ kind: 'range' } & DocxTextRange)
  | { kind: 'table'; story: string; tableIndex: number }
  | { kind: 'control'; story: string; controlId: string }
  | { kind: 'sourcePart'; part: string; partSha256: string; path: number[] };

/** A category of stories a read covers. */
export type DocxStorySelection =
  | 'body'
  | 'headers'
  | 'footers'
  | 'footnotes'
  | 'endnotes'
  | 'comments';

/**
 * A content control's identity and properties. `controlId` is the engine identity, distinct from
 * the authored `w:id` (`ooxmlId`), tag and alias, none of which need be unique.
 */
export interface DocxControlMetadata {
  controlId: string;
  ooxmlId: string | null;
  controlType: string;
  tag: string | null;
  alias: string | null;
  lock: string | null;
  showingPlaceholder: boolean;
  dataBound: boolean;
}

/**
 * A zero-based outline level (0..8). Direct formatting wins, then the paragraph style with
 * inheritance, then document defaults; only without any of them does a `HeadingN` style id count.
 */
export interface DocxHeadingInfo {
  outlineLevel: number;
  source:
    | { kind: 'direct' }
    | { kind: 'style'; styleId: string }
    | { kind: 'documentDefault' }
    | { kind: 'builtinStyleId'; styleId: string };
}

/** A heading paragraph of a story. */
export interface DocxParagraphHeading {
  paraId: string;
  heading: DocxHeadingInfo;
}

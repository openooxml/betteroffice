/**
 * Versioned reads and version-checked edit batches over one DOCX session. Every shape here is
 * plain JSON so it can cross a worker boundary unchanged; the semantics live in Rust.
 *
 * Offsets are half-open UTF-16 positions into a paragraph's projected text. Each inline atom
 * (hard break, image, content control, note reference, field or other embed) occupies one
 * U+FFFC, tabs stay `\t`, and paragraph marks are not part of the text. The accepted view shows
 * pending insertions and hides pending deletions; the original view does the reverse.
 */

import type {
  EditSuccess,
  OperationFailure,
  OperationRefusal,
  ValidationSuccess,
} from '../../../../shared/host-contracts/edits';

export type {
  EditSuccess,
  OperationFailure,
  OperationRefusal,
  ValidationSuccess,
} from '../../../../shared/host-contracts/edits';

export type DocxTextView = 'accepted' | 'original';

/** Provenance only: history and revision authorship are chosen separately. */
export type DocxEditSource = 'host' | 'agent';

/** `separate`: exactly one undo step. `none`: outside undo history, keeping existing entries. */
export type DocxEditHistory = 'separate' | 'none';

export interface DocxParagraphTarget {
  story: string;
  paraId: string;
}

export interface DocxTextPosition {
  paraId: string;
  offset: number;
}

export interface DocxTextRange {
  story: string;
  start: DocxTextPosition;
  end: DocxTextPosition;
  view: DocxTextView;
}

export type DocxSearchScope =
  | { kind: 'story'; story: string }
  | ({ kind: 'paragraph' } & DocxParagraphTarget);

/**
 * The text a step addresses: a paragraph's whole accepted text, an explicit range (both ends in
 * one paragraph), or the one exact, case-sensitive, paragraph-local match of `text`.
 */
export type DocxTextTarget =
  | ({ kind: 'paragraph' } & DocxParagraphTarget)
  | ({ kind: 'range' } & DocxTextRange)
  | { kind: 'search'; text: string; within: DocxSearchScope; view: DocxTextView };

export type DocxEditTarget =
  | DocxTextTarget
  | { kind: 'paragraphs'; story: string; firstParaId: string; lastParaId: string };

/** Refuses the step unless its target currently reads exactly `text`. */
export interface DocxEditGuard {
  text: string;
}

/** Records a text step as a tracked change with host-supplied authorship. */
export interface DocxEditSuggestion {
  author: string;
  date: string;
}

/** One new paragraph; without `styleId` it takes the anchor paragraph's style. */
export interface DocxParagraphInput {
  text: string;
  styleId?: string;
}

export type DocxEditOperation =
  | { op: 'insertText'; target: DocxTextTarget; at: 'start' | 'end'; text: string }
  | { op: 'replaceText'; target: DocxTextTarget; text: string }
  | { op: 'deleteText'; target: DocxTextTarget }
  | {
      op: 'insertParagraphs';
      target: DocxParagraphTarget;
      at: 'start' | 'end';
      paragraphs: readonly DocxParagraphInput[];
    }
  | { op: 'deleteParagraphs'; story: string; firstParaId: string; lastParaId: string }
  | { op: 'setParagraphStyle'; target: DocxParagraphTarget; styleId: string };

export type DocxEditStep = DocxEditOperation & {
  expect?: DocxEditGuard;
  suggest?: DocxEditSuggestion;
};

export interface DocxEditRequest {
  /** The version the targets were read at; a changed document refuses with `stale-version`. */
  expectVersion: string;
  /** Defaults to `host`. */
  source?: DocxEditSource;
  /** Defaults to `separate`. */
  history?: DocxEditHistory;
  steps: readonly DocxEditStep[];
}

/**
 * `locked-target` is locked document content. `read-only` comes only from an editor host that
 * does not accept edits, never from a session.
 */
export type DocxEditFailureCode =
  | 'stale-version'
  | 'missing-target'
  | 'ambiguous-target'
  | 'content-mismatch'
  | 'overlapping-steps'
  | 'locked-target'
  | 'read-only'
  | 'tracked-revision-conflict'
  | 'unsupported'
  | 'invalid-step'
  | 'limit-exceeded';

export type DocxEditFailure = OperationFailure<DocxEditFailureCode, DocxEditTarget>;

export type DocxEditRefusal = OperationRefusal<DocxEditFailure>;

/**
 * What one step did. Ranges use the final accepted view; a deletion reports its collapsed
 * boundary. `removedParagraphs` are historical ids, not anchors.
 */
export interface DocxEditReceipt {
  stepIndex: number;
  changed: boolean;
  range?: DocxTextRange;
  newParagraphs: DocxParagraphTarget[];
  removedParagraphs: DocxParagraphTarget[];
  revisionIds: string[];
}

/** What one step would do; it cannot be committed later. */
export interface DocxEditPreview {
  stepIndex: number;
  target: DocxEditTarget;
  wouldChange: boolean;
  newParagraphCount: number;
  wouldCreateRevisions: boolean;
}

export type DocxEditResult =
  | (EditSuccess<DocxEditReceipt> & { source: DocxEditSource; changedStories: string[] })
  | DocxEditRefusal;

export type DocxValidationResult = ValidationSuccess<DocxEditPreview> | DocxEditRefusal;

export type DocxAtomKind =
  | 'lineBreak'
  | 'image'
  | 'contentControl'
  | 'noteReference'
  | 'field'
  | 'other';

/** One U+FFFC in projected text that stands for an inline atom. */
export interface DocxTextAtom {
  offset: number;
  kind: DocxAtomKind;
}

export interface DocxParagraphText {
  story: string;
  paraId: string;
  text: string;
  styleId?: string;
  atoms: DocxTextAtom[];
}

export interface DocxReadParagraphsRequest {
  /** Defaults to `body`; child stories are never included implicitly. */
  story?: string;
  /** Restricts the read to these paragraphs, returned in story order. */
  paraIds?: readonly string[];
  view: DocxTextView;
}

export type DocxReadParagraphsResult =
  | { ok: true; version: string; view: DocxTextView; paragraphs: DocxParagraphText[] }
  | DocxEditRefusal;

export interface DocxFindTextRequest {
  text: string;
  within: DocxSearchScope;
  view: DocxTextView;
  /** Defaults to 100; at most 10,000. */
  limit?: number;
}

/** One search hit; `range` is reusable as a `{ kind: 'range' }` target. */
export interface DocxTextMatch {
  text: string;
  range: DocxTextRange;
}

export type DocxFindTextResult =
  | { ok: true; version: string; matches: DocxTextMatch[]; truncated: boolean }
  | DocxEditRefusal;

import type { SectionPropertyPatch, SectionStartKind } from './sections';
/**
 * Versioned reads and version-checked edit batches over one DOCX session. Every shape here is
 * plain JSON so it can cross a worker boundary unchanged; the semantics live in Rust.
 *
 * Offsets are half-open UTF-16 positions into a paragraph's projected text. Each inline atom
 * (hard break, image, content control, note reference, field or other embed) occupies one
 * U+FFFC, tabs stay `\t`, and paragraph marks are not part of the text. The accepted view shows
 * pending insertions and hides pending deletions; the original view does the reverse.
 * Reads project paragraph-mark revisions as accepted or rejected, keeping the following
 * paragraph's id on joins; edits still refuse them.
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

import type { DocxContentControlSelector } from './contentControls';
import type { DocxAnchor } from './readTypes';
import type { DocxDrawing } from '../types/content/run';

import type { DocxStoryInfoKind } from './stories';

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
  | { kind: 'paragraphs'; story: string; firstParaId: string; lastParaId: string }
  | { kind: 'contentControl'; selector: DocxContentControlSelector };

/** Refuses the step unless its target currently reads exactly `text`. */
export interface DocxEditGuard {
  text: string;
}

/** Records a text step as a tracked change with host-supplied authorship. */
export interface DocxEditSuggestion {
  author: string;
  date: string;
}

/**
 * Character formatting for one inserted run, written over the formatting plain text gets at the
 * same place; an absent key keeps that formatting. `false`, underline `'none'` and `'noStrike'`
 * turn the mark off, also where a paragraph or character style turns it on. Removing a
 * superscript, subscript or highlight while inserting is not available yet. An unknown key or
 * unsupported value refuses the step.
 */
export interface DocxTextMarks {
  bold?: boolean;
  italic?: boolean;
  /** `true` is `single`; otherwise a `w:u` value such as `double` or `wave`. */
  underline?: boolean | string;
  /** `true` is `sngStrike`. */
  strike?: boolean | 'sngStrike' | 'dblStrike' | 'noStrike';
  verticalAlign?: 'superscript' | 'subscript';
  /** Six hex digits (`RRGGBB`) or a theme colour such as `accent1`. */
  color?: string | { themeColor: string };
  /** One of Word's 16 highlight colour names, such as `yellow` or `darkBlue`. */
  highlight?: string;
  /** 1 to 1638, in half-point steps. */
  fontSizePt?: number;
  fontFamily?: string;
}

/** One run of inserted text. */
export interface DocxTextRun {
  /** Non-empty; no paragraph or line breaks. */
  text: string;
  marks?: DocxTextMarks;
}

/** Plain text, or a non-empty list of formatted runs: exactly one of the two. */
export type DocxTextContent =
  | { text: string; runs?: never }
  | { runs: readonly DocxTextRun[]; text?: never };

/** One new paragraph; without `styleId` it takes the anchor paragraph's style. */
export type DocxParagraphInput = DocxTextContent & { styleId?: string };

export type DocxEditOperation =
  | ({ op: 'insertText'; target: DocxTextTarget; at: 'start' | 'end' } & DocxTextContent)
  | ({ op: 'replaceText'; target: DocxTextTarget } & DocxTextContent)
  | { op: 'deleteText'; target: DocxTextTarget }
  | {
      op: 'insertParagraphs';
      target: DocxParagraphTarget;
      at: 'start' | 'end';
      paragraphs: readonly DocxParagraphInput[];
    }
  | { op: 'deleteParagraphs'; story: string; firstParaId: string; lastParaId: string }
  | { op: 'insertSectionBreakAfter'; target: DocxParagraphTarget; kind: SectionStartKind }
  | { op: 'setSectionProperties'; section: string; properties: SectionPropertyPatch }
  | { op: 'setParagraphStyle'; target: DocxParagraphTarget; styleId: string }
  | Omit<DocxSetContentControlTextStep, 'expect'>;

/**
 * Replaces a plain- or rich-text content control's content with plain text and clears its
 * placeholder state. CRLF becomes LF; LF is a line break (a paragraph break in a rich-text block
 * control), and a plain-text control accepts it only with `w:multiLine`. The text takes the
 * formatting of the control's first text run (a control showing its placeholder takes its own run
 * properties). `expect` compares against the control's `value`; `suggest` is refused.
 */
export interface DocxSetContentControlTextStep {
  op: 'setContentControlText';
  target: DocxContentControlSelector;
  text: string;
  expect?: DocxEditGuard;
}

export type DocxEditStep =
  | (Exclude<DocxEditOperation, { op: 'setContentControlText' }> & {
      expect?: DocxEditGuard;
      suggest?: DocxEditSuggestion;
    })
  | DocxSetContentControlTextStep;

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

/** Why a step was refused, beside its failure code: a `setContentControlText` step, or a step aimed at a table cell that has no story. */
export type DocxEditFailureReason =
  | 'missing-control'
  | 'missing-tag'
  | 'ambiguous-tag'
  | 'ambiguous-control-id'
  | 'missing-ooxml-id'
  | 'ambiguous-ooxml-id'
  | 'content-locked'
  | 'bound-control'
  | 'unsupported-control-type'
  | 'unsupported-children'
  | 'nested-controls'
  | 'unknown-lock'
  | 'unsupported-suggestion'
  | 'provenance-unavailable'
  | 'unsupported-story'
  | 'multiline-not-allowed'
  | 'invalid-text'
  | 'cell-has-no-story';

export type DocxEditFailure = OperationFailure<DocxEditFailureCode, DocxEditTarget> & {
  reason?: DocxEditFailureReason;
};

/** The content control a step resolved to, and where it is. */
export interface DocxResolvedControl {
  controlId: string;
  anchor: DocxAnchor;
}

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
  control?: DocxResolvedControl;
}

/** What one step would do; it cannot be committed later. */
export interface DocxEditPreview {
  stepIndex: number;
  target: DocxEditTarget;
  wouldChange: boolean;
  newParagraphCount: number;
  wouldCreateRevisions: boolean;
  control?: DocxResolvedControl;
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

export type { DocxDrawing } from '../types/content/run';

/** One U+FFFC in projected text that stands for an inline atom. */
export interface DocxTextAtom {
  offset: number;
  kind: DocxAtomKind;
  /** The drawing's `wp:docPr` properties, on atoms whose embed carries them. */
  drawing?: DocxDrawing;
  /** Drawings inside content-control atoms in the requested view;
   * drawings inside field results are not listed. */
  drawings?: DocxDrawing[];
}

/** Numbering in the requested view, counted from the start of the root story. */
export interface DocxParagraphList {
  /** Rendered marker, including when hidden; null when unresolved, empty string when empty. */
  marker: string | null;
  /** Zero-based list level, 0..8. */
  level: number;
  /** OOXML numbering instance id. */
  numId: string;
  /** OOXML numFmt, including `bullet`. */
  format: string;
}

export interface DocxParagraphText {
  story: string;
  paraId: string;
  text: string;
  styleId?: string;
  /** Style-resolved outline level in the requested view, 0..8, or null for body text. */
  outlineLevel?: number | null;
  /** Numbering in the requested view, counted from the root story start; null without numbering. */
  list?: DocxParagraphList | null;
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

export interface DocxReadStoriesRequest {
  /**
   * `'all'` (the default) reads every story in id order. Listed entries are read in their order
   * without repeats; an entry naming a kind stands for its stories in id order.
   */
  stories?: 'all' | ReadonlyArray<DocxStoryInfoKind | (string & {})>;
  /** Matches listed kinds against each story's root, so `header` also selects header cells. */
  byRoot?: boolean;
  view: DocxTextView;
  /** Refuses with `stale-version` when the document is at another version. */
  expectVersion?: string;
}

/** One story's paragraphs, or why it could not be read. */
export type DocxStoryText =
  | { story: string; ok: true; paragraphs: DocxParagraphText[] }
  | { story: string; ok: false; failure: DocxEditFailure };

export type DocxReadStoriesResult =
  | {
      ok: true;
      version: string;
      view: DocxTextView;
      stories: DocxStoryText[];
      /**
   * The read limits stopped the read before the next selected story. A story larger than the
   * limits on its own reports `limit-exceeded` instead; page it with `readParagraphs` `paraIds`.
   */
      truncated: boolean;
    }
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

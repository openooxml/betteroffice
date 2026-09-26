/**
 * Versioned reads and version-checked edit batches over one presentation session. Every shape
 * here is plain JSON so it can cross a worker boundary unchanged; the semantics live in Rust.
 *
 * Text offsets are half-open, story-local UTF-16 positions. A story reads as its paragraphs'
 * text joined by `\n` with no trailing separator; soft line breaks read as `\n` too, and
 * `PptxParagraphText` tells them apart. Slide, shape, story and paragraph ids anchor targets
 * within one session; they are not promised to survive save and reopen.
 */

import type {
  EditSuccess,
  OperationFailure,
  OperationRefusal,
  ValidationSuccess,
} from '../../../shared/host-contracts/edits';
import type {
  ParagraphAlignment,
  ShapeFill,
  ShapeOutline,
  ShapeRect,
  ShapeStroke,
  SlideSnapshot,
  TextStylePatch,
} from './types';

export type {
  EditSuccess,
  OperationFailure,
  OperationRefusal,
  ValidationSuccess,
} from '../../../shared/host-contracts/edits';

/** Provenance only: it selects neither permissions nor history. */
export type PptxEditSource = 'host' | 'agent';

/** `separate`: exactly one undo step. `none`: outside undo history, keeping existing entries. */
export type PptxEditHistory = 'separate' | 'none';

export interface PptxSlideTarget {
  slideId: string;
}

/** A shape addressed through the slide that owns it. */
export interface PptxShapeTarget extends PptxSlideTarget {
  shapeId: string;
}

/**
 * A text story addressed through its slide and owning shape. Targets reject unknown fields, so
 * copy the ids off a read record rather than passing the record itself.
 */
export interface PptxStoryTarget extends PptxShapeTarget {
  storyId: string;
}

export interface PptxTextRange extends PptxStoryTarget {
  start: number;
  end: number;
}

/** An explicit range, or the one exact, case-sensitive, paragraph-local match of `text`. */
export type PptxTextTarget =
  | ({ kind: 'range' } & PptxTextRange)
  | { kind: 'search'; within: PptxStoryTarget; text: string };

/** Refuses the step unless its target currently reads exactly `text`. */
export interface PptxTextGuard {
  text: string;
}

/**
 * One batch step. Insertion, replacement and deletion stay inside one paragraph; formatting and
 * alignment may span paragraphs of one story. Geometry, fill and outline steps address shapes at
 * the top of a slide, and fill and outline only preset shapes.
 */
export type PptxEditStep =
  | {
      op: 'insertText';
      target: PptxTextTarget;
      at: 'start' | 'end';
      text: string;
      expect?: PptxTextGuard;
    }
  | { op: 'replaceText'; target: PptxTextTarget; text: string; expect?: PptxTextGuard }
  | { op: 'deleteText'; target: PptxTextTarget; expect?: PptxTextGuard }
  | { op: 'formatText'; target: PptxTextTarget; patch: TextStylePatch; expect?: PptxTextGuard }
  | {
      op: 'setParagraphAlignment';
      target: PptxTextTarget;
      /** `null` restores the inherited alignment. */
      alignment: ParagraphAlignment | null;
      expect?: PptxTextGuard;
    }
  | { op: 'setSlideNotes'; target: PptxSlideTarget; text: string; expect?: PptxTextGuard }
  | { op: 'setShapeRect'; target: PptxShapeTarget; rect: ShapeRect; expect?: { rect: ShapeRect } }
  | {
      op: 'setShapeFill';
      target: PptxShapeTarget;
      /** A `#RRGGBB` solid fill; `null` removes the fill. */
      color: string | null;
      /** The authored fill, as `snapshot()` reports it. */
      expect?: { fill: ShapeFill | null };
    }
  | {
      op: 'setShapeStroke';
      target: PptxShapeTarget;
      stroke: ShapeStroke;
      /** The authored outline, as `snapshot()` reports it. */
      expect?: { outline: ShapeOutline | null };
    };

export interface PptxEditRequest {
  /** The version the targets were read at; a changed deck refuses with `stale-version`. */
  expectVersion: string;
  /** Defaults to `host`. */
  source?: PptxEditSource;
  /** Defaults to `separate`. */
  history?: PptxEditHistory;
  steps: readonly PptxEditStep[];
}

/** `read-only` comes only from an editor host that does not accept edits, never from a session. */
export type PptxEditFailureCode =
  | 'stale-version'
  | 'missing-target'
  | 'ambiguous-target'
  | 'content-mismatch'
  | 'overlapping-steps'
  | 'unsupported'
  | 'invalid-step'
  | 'limit-exceeded'
  | 'read-only';

export type PptxEditTarget =
  | PptxTextTarget
  | ({ kind: 'story' } & PptxStoryTarget)
  | ({ kind: 'shape' } & PptxShapeTarget)
  | ({ kind: 'slide' } & PptxSlideTarget);

export type PptxEditFailure = OperationFailure<PptxEditFailureCode, PptxEditTarget>;

export type PptxEditRefusal = OperationRefusal<PptxEditFailure>;

/** What one step did. Text targets describe the final state; a deletion reports its collapsed boundary. */
export interface PptxEditReceipt {
  stepIndex: number;
  changed: boolean;
  target: PptxEditTarget;
}

/** What one step would do at its resolved pre-batch target; it cannot be committed later. */
export interface PptxEditPreview {
  stepIndex: number;
  target: PptxEditTarget;
  wouldChange: boolean;
}

export type PptxEditResult =
  | (EditSuccess<PptxEditReceipt> & {
      source: PptxEditSource;
      changedSlides: string[];
      changedStories: string[];
    })
  | PptxEditRefusal;

export type PptxValidationResult = ValidationSuccess<PptxEditPreview> | PptxEditRefusal;

export interface PptxReadRequest {
  /** Restricts the read to these slides, returned in deck order; every slide when omitted. */
  slideIds?: readonly string[];
}

/** The story offsets of a field result, such as a slide number. */
export interface PptxTextField {
  start: number;
  end: number;
  fieldType?: string;
}

export interface PptxParagraphText {
  paragraphId: string;
  /** Story offsets of the paragraph's text; a following separator sits at `end`. */
  start: number;
  end: number;
  /** Story offsets of soft line breaks, which read as `\n` inside the paragraph. */
  lineBreaks: number[];
  /** Field results that text steps leave whole. */
  fields: PptxTextField[];
  /**
   * False when saving could not keep a field once the paragraph changes: the field is empty or
   * left the paragraph's unchanged leading or trailing text. Steps changing it refuse.
   */
  editable: boolean;
}

export interface PptxStoryText extends PptxStoryTarget {
  /** Paragraph texts joined by `\n`. */
  text: string;
  paragraphs: PptxParagraphText[];
}

export type PptxReadResult =
  | { ok: true; version: string; slides: SlideSnapshot[]; stories: PptxStoryText[] }
  | PptxEditRefusal;

/** A slide, one shape and its descendants, or one story; `storyId` needs `shapeId`. */
export interface PptxFindScope {
  slideId: string;
  shapeId?: string;
  storyId?: string;
}

export interface PptxFindRequest {
  text: string;
  /** The whole deck when omitted. */
  within?: PptxFindScope;
  /** Defaults to 100; at most 10,000. */
  limit?: number;
}

/** One search hit; `range` is reusable as a `{ kind: 'range' }` target. */
export interface PptxFindMatch {
  text: string;
  range: PptxTextRange;
}

export type PptxFindResult =
  | { ok: true; version: string; matches: PptxFindMatch[]; truncated: boolean }
  | PptxEditRefusal;

import type {
  DeckSnapshot,
  ParagraphAlignment,
  ShapeDraft,
  PresetShapeDraft,
  PictureDraft,
  ShapeRect,
  ShapeSnapshot,
  ShapeStroke,
  SlideDisplayList,
  TextParagraphDraft,
  TextStyleSnapshot,
} from './types';

export type ProposalEdit =
  | { type: 'addTextBox'; ref?: string; slideId: string; draft: ShapeDraft }
  | { type: 'addShape'; ref?: string; slideId: string; draft: PresetShapeDraft }
  | { type: 'addPicture'; ref?: string; slideId: string; draft: PictureDraft }
  | { type: 'removeShape'; slideId: string; shapeId: string }
  | { type: 'insertSlide'; ref?: string; afterSlideId: string | null; layoutPartPath?: string }
  | { type: 'duplicateSlide'; ref?: string; slideId: string; afterSlideId?: string | null }
  | { type: 'deleteSlide'; slideId: string }
  | { type: 'moveSlide'; slideId: string; afterSlideId: string | null }
  | {
      type: 'replaceText';
      storyId: string;
      start: number;
      end: number;
      text: string;
      style?: Partial<TextStyleSnapshot> | null;
    }
  | {
      type: 'formatText';
      storyId: string;
      start: number;
      end: number;
      patch: Partial<TextStyleSnapshot>;
    }
  | {
      type: 'setParagraphAlignment';
      storyId: string;
      start: number;
      end: number;
      alignment: ParagraphAlignment | 'justLow' | 'dist' | 'thaiDist' | null;
    }
  | { type: 'setStoryParagraphs'; storyId: string; paragraphs: TextParagraphDraft[] }
  | { type: 'setShapeRect'; slideId: string; shapeId: string; rect: ShapeRect }
  | {
      type: 'setShapeFill';
      slideId: string;
      shapeId: string;
      color: string | null;
    }
  | {
      type: 'setShapeStroke';
      slideId: string;
      shapeId: string;
      stroke: { [K in keyof ShapeStroke]: ShapeStroke[K] | null };
    }
  | {
      type: 'setShapeAdjust';
      slideId: string;
      shapeId: string;
      adjustments: Record<string, number>;
    }
  | { type: 'setSlideNotes'; slideId: string; text: string };

export interface ProposalStripEntry {
  slideId: string;
  status: 'unchanged' | 'changed' | 'inserted' | 'deleted' | 'moved';
  liveIndex: number | null;
  previewIndex: number | null;
}

export interface ProposalChange {
  kind?: 'content' | 'inserted' | 'deleted' | 'moved';
  beforeIndex?: number | null;
  afterIndex?: number | null;
  slideId: string;
  shapeId: string | null;
  before: ShapeSnapshot | null;
  after: ShapeSnapshot | null;
  oldText: string;
  newText: string;
}

export interface Proposal {
  id: string;
  agentId: string;
  note: string | null;
  edits: ProposalEdit[];
  changes: ProposalChange[];
  staleTargets: string[];
  refs?: Record<string, string>;
}

export interface ProposalPreview {
  proposal: Proposal;
  snapshot: DeckSnapshot;
}

export interface ProposalTextChange {
  storyId: string;
  start: number;
  end: number;
  kind: 'insertion' | 'deletion';
}

export interface ProposalShapeChange {
  shapeId: string;
  kind: 'inserted' | 'deleted' | 'moved' | 'resized' | 'changed' | 'reordered';
}

export interface ProposalDiffSlide extends ProposalPreview {
  shapeChanges?: ProposalShapeChange[];
  frame: SlideDisplayList;
  textChanges: ProposalTextChange[];
  media?: { assetId: string; bytes: ArrayBuffer }[];
}

export interface ProposalAcceptance {
  proposalId: string;
  applied: boolean;
  snapshot: DeckSnapshot;
}

export class StaleProposalError extends Error {
  readonly targets: string[];

  constructor(targets: string[]) {
    super(`Proposal targets changed: ${targets.join(', ')}`);
    this.name = 'StaleProposalError';
    this.targets = targets;
  }
}

export class ProposalPictureMediaTooLargeError extends Error {
  readonly code = 'proposalPictureMediaTooLarge';

  constructor(readonly totalBytes: number, readonly limitBytes = 8 * 1024 * 1024) {
    super(`Proposal picture media exceeds ${limitBytes} decoded bytes (${totalBytes} bytes)`);
    this.name = 'ProposalPictureMediaTooLargeError';
  }
}

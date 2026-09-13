import type {
  DeckSnapshot,
  ParagraphAlignment,
  ShapeRect,
  ShapeSnapshot,
  ShapeStroke,
  SlideDisplayList,
  TextStyleSnapshot,
} from './types';

export type ProposalEdit =
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

export interface ProposalChange {
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

export interface ProposalDiffSlide extends ProposalPreview {
  frame: SlideDisplayList;
  textChanges: ProposalTextChange[];
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

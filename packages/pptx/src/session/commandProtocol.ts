import type { SessionFailure } from '../../../../shared/office-session/types';
import type { PresentationHandle } from '../wasm/loader';
import type { PptxEditRequest, PptxEditResult, PptxReadRequest, PptxReadResult, PptxStoryTarget } from '../edits';
import type { DeckSnapshot, ParagraphAlignment, ShapeRect, ShapeDraft, PresetShapeDraft, ShapeReceipt, TextStyle, TextStyleSnapshot } from '../types';
import type {
  PresentationSessionOpenOptions, PresentationSlideSummary, PresentationWireFrame,
} from './methods';

export type PresentationCommandAccessMethods = Pick<PresentationHandle,
  'snapshot' | 'story' | 'findText' | 'validateEdits' | 'exportStructured' | 'exportMarkdown' | 'searchText' | 'listProposals' | 'comments' | 'previewProposal' | 'layoutProposalSlide' | 'layoutProposalDiffSlide' | 'layoutSlide' | 'encodeStateVector' | 'encodeStateAsUpdate'>;
export type PresentationCommandAccessMethod = keyof PresentationCommandAccessMethods;
export type PresentationCommandAccessResult = ReturnType<PresentationCommandAccessMethods[PresentationCommandAccessMethod]>;

export type PresentationCommandMutationMethods = Pick<PresentationHandle,
  'setShapeFill' | 'setShapeStroke' | 'setShapeAdjust' | 'removeShape' | 'bringShapeToFront' |
  'sendShapeToBack' | 'bringShapeForward' | 'sendShapeBackward' | 'insertSlide' | 'deleteSlide' | 'moveSlide' | 'setSlideNotes' | 'addComment' | 'replyToComment' | 'setCommentStatus' | 'setCommentPosition' | 'removeComment' | 'setCommentFlavor' | 'propose' | 'acceptProposal' | 'rejectProposal'>;
export const PRESENTATION_COMMAND_MUTATIONS = ['setShapeFill', 'setShapeStroke', 'setShapeAdjust', 'removeShape',
  'bringShapeToFront', 'sendShapeToBack', 'bringShapeForward', 'sendShapeBackward', 'insertSlide', 'deleteSlide', 'moveSlide', 'setSlideNotes', 'addComment', 'replyToComment', 'setCommentStatus', 'setCommentPosition', 'removeComment', 'setCommentFlavor', 'propose', 'acceptProposal', 'rejectProposal'] as const satisfies readonly (keyof PresentationCommandMutationMethods)[];

export type PresentationCommandMutation = { [K in keyof PresentationCommandMutationMethods]: {
  method: K; args: Parameters<PresentationCommandMutationMethods[K]>;
} }[keyof PresentationCommandMutationMethods];

export interface PresentationCommandIdentity {
  session: string;
  incarnation: string;
}

export interface PresentationCommandBasis {
  epoch: number;
  version: string;
}

export interface PresentationCommandSelection extends PptxStoryTarget {
  kind?: 'shape';
  anchor: number;
  focus: number;
  affinity: 'upstream' | 'downstream';
  goalX?: number;
}

export interface PresentationCommandObjectSelection extends Omit<PresentationCommandSelection, 'kind'> {
  kind: 'object';
  token: string;
  bounds: ShapeRect;
  rect: ShapeRect | null;
  canMove: boolean;
  canResize: boolean;
  units: { x: number; y: number };
}

export type PresentationCommandSelected = PresentationCommandSelection | PresentationCommandObjectSelection;

export interface PresentationCommandObjectTarget {
  slideId: string;
  shapeId: string;
  token: string;
}

export type PresentationCommandSelectionReference =
  | { source: 'confirmed'; selection: PresentationCommandSelection }
  | { source: 'afterSeq'; seq: number };

export type PresentationCommandMotion = 'left' | 'right' | 'wordLeft' | 'wordRight'
  | 'lineStart' | 'lineEnd' | 'storyStart' | 'storyEnd' | 'up' | 'down';

export type PresentationCommandFrameTag = Pick<PresentationCommandFrame,
  'slideId' | 'slideRevision' | 'coversThroughSeq' | 'documentEpoch' | 'incarnation'>;

export interface PresentationCommandInput {
  inputTime: number;
  undoGroup: string;
}

export interface PresentationCommandTextInput extends PresentationCommandInput {
  at: PresentationCommandSelectionReference;
}

export type PresentationCommandPayloads = {
  setChangeObserver: PresentationCommandInput & { enabled: boolean };
  invoke: PresentationCommandInput & PresentationCommandMutation;
  applyRemoteUpdate: PresentationCommandInput & { bytes: ArrayBuffer };
  applyEdits: PresentationCommandInput & { request: PptxEditRequest };
  registerFont: PresentationCommandInput & { font: { family: string; bytes: ArrayBuffer; bold?: boolean; italic?: boolean }; fallback?: boolean };
  insertTextBox: PresentationCommandInput & { slideId: string; draft: ShapeDraft };
  insertShape: PresentationCommandInput & { slideId: string; draft: PresetShapeDraft };
  insertPicture: PresentationCommandInput & {
    slideId: string; name: string; rect: ShapeRect; contentType: string; bytes: ArrayBuffer;
  };
  selectObject: PresentationCommandInput & {
    frame: PresentationCommandFrameTag;
  } & ({ x: number; y: number; scale: number } | { shapeId: string });
  moveShape: PresentationCommandInput & {
    frame: PresentationCommandFrameTag;
  } & ({ target: PresentationCommandObjectTarget; x: number; y: number }
    | { at: PresentationCommandSelectionReference; shapeId: string; delta: { x: number; y: number } });
  setShapeRect: PresentationCommandInput & {
    frame: PresentationCommandFrameTag; target: PresentationCommandObjectTarget; rect: ShapeRect;
  };
  undo: PresentationCommandInput;
  redo: PresentationCommandInput;
  activateSlide: Partial<PresentationCommandInput> & { slideId: string };
  insertText: PresentationCommandTextInput & { text: string; style?: TextStyle };
  replaceText: PresentationCommandTextInput & { text: string; style?: TextStyle };
  formatText: PresentationCommandTextInput & { patch: TextStyle; toggle?: 'bold' | 'italic' | 'underline';
    fontSizeStep?: { direction: 'increase' | 'decrease'; sizes: readonly number[] } };
  setParagraphAlignment: PresentationCommandTextInput & { alignment: ParagraphAlignment | null };
  deleteBackward: PresentationCommandTextInput & { unit: 'codePoint'; count?: number };
  deleteForward: PresentationCommandTextInput & { unit: 'codePoint'; count?: number };
  paragraphBreak: PresentationCommandTextInput;
  moveSelection: PresentationCommandTextInput & { motion: PresentationCommandMotion; extend: boolean };
  selectAtPoint: PresentationCommandInput & {
    at?: PresentationCommandSelectionReference;
    frame: PresentationCommandFrameTag; x: number; y: number; extend: boolean;
    granularity?: 'character' | 'word' | 'paragraph'; drag?: boolean;
  };
  setSelection: PresentationCommandInput & {
    selection: PresentationCommandSelection | null; at?: PresentationCommandSelectionReference;
  };
  compositionStart: PresentationCommandTextInput & { compositionId: string };
  compositionEnd: PresentationCommandInput & {
    compositionId: string; cancelled?: boolean; at?: PresentationCommandSelectionReference;
  };
  compositionCommit: PresentationCommandTextInput & { compositionId: string; text: string; style?: TextStyle };
  setReadOnly: PresentationCommandInput & { readOnly: boolean };
};

export type PresentationCommandKind = keyof PresentationCommandPayloads;

export type PresentationCommandAdmission = 'view' | 'read' | 'local-write' | 'remote-update' | 'mode';

export const PRESENTATION_COMMAND_ADMISSION = Object.freeze({
  setChangeObserver: 'mode', invoke: 'local-write',
  applyRemoteUpdate: 'remote-update',
  applyEdits: 'local-write', registerFont: 'local-write',
  insertTextBox: 'local-write', insertShape: 'local-write', insertPicture: 'local-write',
  moveShape: 'local-write', setShapeRect: 'local-write', undo: 'local-write', redo: 'local-write',
  insertText: 'local-write', replaceText: 'local-write', formatText: 'local-write',
  setParagraphAlignment: 'local-write', deleteBackward: 'local-write', deleteForward: 'local-write',
  paragraphBreak: 'local-write', compositionStart: 'local-write', compositionCommit: 'local-write',
  activateSlide: 'view', selectObject: 'view', moveSelection: 'view', selectAtPoint: 'view',
  setSelection: 'view', compositionEnd: 'view', setReadOnly: 'mode',
} as const satisfies Record<PresentationCommandKind, PresentationCommandAdmission>);

export type PresentationCommandIntent = {
  [K in PresentationCommandKind]: { kind: K; payload: PresentationCommandPayloads[K] }
}[PresentationCommandKind];

export type PresentationCommandEnvelope = PresentationCommandIdentity & {
  seq: number;
  basis: PresentationCommandBasis;
  afterSeq?: number;
} & PresentationCommandIntent;

export type PresentationCommandRefusalReason =
  | 'bad-target'
  | 'invalid-offset'
  | 'stale-basis'
  | 'read-only'
  | 'dependency-refused'
  | 'invalid-composition'
  | 'invalid-payload'
  | 'unsupported'
  | 'unknown-outcome';

export class CommandOutcomeUnknown extends Error {
  readonly refusal: { seq: number; reason: 'unknown-outcome' };

  constructor(readonly seq: number, message: string) {
    super(message);
    this.name = 'CommandOutcomeUnknown';
    this.refusal = { seq, reason: 'unknown-outcome' };
  }
}

export type PresentationCommandFailure = CommandOutcomeUnknown | SessionFailure;

export interface PresentationCommandOutcome {
  seq: number;
  status: 'applied' | 'noop' | 'refused' | 'cancelled';
  reason?: PresentationCommandRefusalReason;
  committed: boolean;
}

export interface PresentationCommandToolbarContext {
  target: PptxStoryTarget | null;
  textContext?: 'caret' | 'range' | 'shape' | null;
  storedMarks: TextStyle;
  formatting: TextStyleSnapshot | null;
  alignment: ParagraphAlignment | null;
  hyperlink: { url: string } | null;
  table: { shapeId: string; row: number; column: number } | null;
  canEditText: boolean;
  echoRun?: 'before' | 'after' | null;
}

export interface PresentationCommandEditState {
  readOnly: boolean;
  notes?: { slideId: string; text: string } | null;
  proposals?: import('../proposals').Proposal[];
  object?: { formattable: boolean; geometry: string | null; fill: string | null; stroke: string | null;
    strokeWidth: number | null; adjustments: Record<string, number>; order: { index: number; count: number } | null } | null;
  canUndo: boolean;
  canRedo: boolean;
  toolbar: PresentationCommandToolbarContext;
}

export interface PresentationCommandFrame extends PresentationWireFrame {
  slideId: string;
  slideRevision: number;
  coversThroughSeq: number;
  documentEpoch: number;
  incarnation: string;
}

export interface PresentationCommandAck extends PresentationCommandIdentity {
  epoch: number;
  throughSeq: number;
  version: string;
  selection: PresentationCommandSelected | null;
  editState: PresentationCommandEditState;
  outcomes: PresentationCommandOutcome[];
  dirtySlideIds: string[];
  mutationResult?: unknown;
  change?: PresentationCommandIdentity & { throughSeq: number; version: string; snapshot: DeckSnapshot };
  updates?: { bytes: ArrayBuffer; origin: 'local' | 'remote' }[];
  creation?: ShapeReceipt;
  editResult?: PptxEditResult;
  fontId?: number;
  activeSlideId?: string | null;
  historyThroughSeq?: number;
  slides?: PresentationSlideSummary[];
  size?: { width: number; height: number };
  frame?: PresentationCommandFrame | null;
}

export interface PresentationCommandOpenOptions
  extends PresentationSessionOpenOptions, PresentationCommandIdentity {
  readOnly?: boolean;
  activeSlideIndex?: number;
  collaboration?: boolean;
  observeChanges?: boolean;
}

export interface PresentationCommandOpenResult extends PresentationCommandAck {
  recoveryUpdate?: ArrayBuffer;
  slides: PresentationSlideSummary[];
  size: { width: number; height: number };
  activeSlideId: string | null;
  frame: PresentationCommandFrame | null;
}

export interface PresentationCommandFence extends PresentationCommandIdentity {
  seq: number;
}

export interface PresentationCommandThumbnailRequest extends PresentationCommandIdentity {
  slideId: string;
  knownRevision?: number;
}

export type PresentationCommandThumbnailResult = PresentationCommandIdentity & {
  epoch: number;
  slideId: string;
  slideRevision: number;
} & (
  | { status: 'unchanged' }
  | { status: 'refused'; reason: 'bad-target' }
  | { status: 'frame'; frame: PresentationCommandFrame }
);

export interface PresentationCommandSaveResult extends PresentationCommandIdentity {
  epoch: number;
  throughSeq: number;
  version: string;
  bytes: ArrayBuffer;
  recoveryUpdate?: ArrayBuffer;
}

export interface PresentationCommandReadResult extends PresentationCommandAck {
  result: PptxReadResult;
}

export type PresentationCommandMethods = {
  openCommandEditor(bytes: ArrayBuffer, options: PresentationCommandOpenOptions): PresentationCommandOpenResult;
  command(envelope: PresentationCommandEnvelope): PresentationCommandAck;
  thumbnailFrame(request: PresentationCommandThumbnailRequest): PresentationCommandThumbnailResult;
  flushCommandEditor(fence: PresentationCommandFence): PresentationCommandAck;
  saveCommandEditor(fence: PresentationCommandFence): PresentationCommandSaveResult;
  accessCommandEditor(fence: PresentationCommandFence, method: PresentationCommandAccessMethod, args: unknown[]): PresentationCommandAccessResult;
  checkpointCommandEditor(fence: PresentationCommandFence): PresentationCommandSaveResult;
  prepareCommandPicture(fence: PresentationCommandFence, bytes: ArrayBuffer): ArrayBuffer;
  readEditor(fence: PresentationCommandFence, request?: PptxReadRequest): PresentationCommandReadResult;
};

export type PresentationCommandEvents = {
  ack: PresentationCommandAck;
};

export interface PresentationCommandClient {
  readonly state: PresentationCommandOpenResult;
  dispatch(intent: PresentationCommandIntent, basis?: PresentationCommandBasis): Promise<PresentationCommandAck>;
  flush(seq?: number): Promise<PresentationCommandAck>;
  save(): Promise<Uint8Array>;
  access?<K extends PresentationCommandAccessMethod>(method: K, ...args: Parameters<PresentationCommandAccessMethods[K]>): Promise<ReturnType<PresentationCommandAccessMethods[K]>>;
  checkpoint?(): Promise<PresentationCommandSaveResult>;
  preparePicture?(bytes: ArrayBuffer): Promise<ArrayBuffer>;
  read(request?: PptxReadRequest): Promise<PresentationCommandReadResult>;
  thumbnail(slideId: string, knownRevision?: number): Promise<PresentationCommandThumbnailResult>;
  subscribe(listener: (ack: PresentationCommandAck) => void): () => void;
  subscribeFrame(listener: (frame: PresentationCommandFrame) => void): () => void;
  subscribeMedia(listener: (media: PresentationCommandFrame['media']) => void): () => void;
  subscribeUpdate?(listener: (update: Uint8Array, origin: 'local' | 'remote') => void): () => void;
  subscribeFailure(listener: (error: PresentationCommandFailure) => void): () => void;
  dispose(): Promise<void>;
}

export const PRESENTATION_COMMAND_METHODS = {
  openCommandEditor: true,
  command: true,
  thumbnailFrame: true,
  flushCommandEditor: true,
  saveCommandEditor: true,
  readEditor: true,
  prepareCommandPicture: true,
  accessCommandEditor: true,
  checkpointCommandEditor: true,
} satisfies { readonly [K in keyof PresentationCommandMethods]-?: true };

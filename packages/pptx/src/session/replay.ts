import type { MethodPolicies } from '../../../../shared/office-session/types';
import type { SessionClient } from '../../../../shared/office-session/client';
import type { PresentationEditorConnection } from './client';
import type {
  PresentationReplayEnvelope, PresentationReplayOp, PresentationReplayReply,
} from '../wasm/loader';
import type {
  PresentationFrame, PresentationSessionOpenOptions, PresentationSessionState, PresentationWireFrame,
} from './methods';

export type PptxWorkerEditorOperation = PresentationReplayOp;
export type PptxWorkerEditorReply = PresentationReplayReply;
export type PptxWorkerEditorFrame = PresentationFrame & { slideId: string; sequence: number };
export type EditorPosition = { version: string; sequence: number };
export type EditorBaseline = EditorPosition & {
  state: PresentationSessionState;
  hydration: string;
  module: WebAssembly.Module;
};
export type EditorReplayReply = PresentationReplayReply & { projection: PresentationSessionState };
export type PresentationEditorMethods = {
  beginEditor(bytes: ArrayBuffer, options: PresentationSessionOpenOptions): EditorBaseline;
  attachPeer(version: string, sequence: number): EditorPosition;
  replay(envelope: PresentationReplayEnvelope): EditorReplayReply;
  flush(sequence: number, version: string): EditorPosition;
  editorFrame(slideId: string): PresentationWireFrame & { slideId: string; sequence: number };
  editorSave(sequence: number, version: string): ArrayBuffer;
};

export const presentationEditorSessionInternals = new WeakMap<PresentationEditorConnection,
  { client: SessionClient<PresentationEditorMethods, {}>; peerAttached: boolean }>();

export const PRESENTATION_EDITOR_METHODS = {
  beginEditor: true, attachPeer: true, replay: true, flush: true, editorFrame: true, editorSave: true,
} satisfies { [K in keyof PresentationEditorMethods]: true };

export const PRESENTATION_EDITOR_POLICIES: MethodPolicies<PresentationEditorMethods> = {
  beginEditor: { lane: 'interactive' },
  attachPeer: { lane: 'interactive' },
  replay: { lane: 'input', mutates: true, userInput: true },
  flush: { lane: 'interactive' },
  editorFrame: { lane: 'interactive', reframes: true },
  editorSave: { lane: 'interactive' },
};

export const PRESENTATION_REPLAY_MUTATORS = {
  insertText: true, deleteText: true, formatText: true, insertParagraphBreak: true,
  setParagraphAlignment: true, insertSlide: true, deleteSlide: true, moveSlide: true,
  setSlideNotes: true, addTextBox: true, addShape: true, addPicture: true, removeShape: true,
  moveShape: true, resizeShape: true, setShapeRect: true, setShapeFill: true, setShapeStroke: true,
  setShapeAdjust: true, bringShapeToFront: true, sendShapeToBack: true, bringShapeForward: true,
  sendShapeBackward: true, addComment: true, replyToComment: true, setCommentStatus: true,
  setCommentPosition: true, removeComment: true, setCommentFlavor: true, propose: true,
  acceptProposal: true, rejectProposal: true, applyEdits: true, addUndoBoundary: true,
  undo: true, redo: true,
} satisfies { [K in PresentationReplayOp['method']]: true };

export function sameReplayValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false;
  if (Array.isArray(left) !== Array.isArray(right)) return false;
  const a = Object.keys(left);
  const b = Object.keys(right);
  return a.length === b.length && a.every((key) => Object.prototype.hasOwnProperty.call(right, key) &&
    sameReplayValue((left as Record<string, unknown>)[key], (right as Record<string, unknown>)[key]));
}

export function matchingReplayReply(local: PresentationReplayReply, remote: PresentationReplayReply): boolean {
  return remote.consumed === true && local.sequence === remote.sequence &&
    local.revision === remote.revision && local.version === remote.version &&
    local.engineVersion === remote.engineVersion && sameReplayValue(local.outcome, remote.outcome);
}

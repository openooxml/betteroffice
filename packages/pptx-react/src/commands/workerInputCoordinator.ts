import { PptxPeerNotReadyError, PptxWorkerEditorDisposedError } from '@betteroffice/pptx';
import type { PptxWorkerEditorAccess, PptxWorkerEditorOperation, PptxWorkerEditorSession } from '@betteroffice/pptx';
import type { PptxInputCoordinator } from './inputCoordinator';
import { PptxCommandAdmissionError } from './createPptxCommandStore';

const mutations = {
  insertText: true, deleteText: true, formatText: true, insertParagraphBreak: true,
  setParagraphAlignment: true, insertSlide: true, deleteSlide: true, moveSlide: true,
  setSlideNotes: true, addTextBox: true, addShape: true, addPicture: true, removeShape: true,
  moveShape: true, resizeShape: true, setShapeRect: true, setShapeFill: true, setShapeStroke: true,
  setShapeAdjust: true, bringShapeToFront: true, sendShapeToBack: true, bringShapeForward: true,
  sendShapeBackward: true, addComment: true, replyToComment: true, setCommentStatus: true,
  setCommentPosition: true, removeComment: true, setCommentFlavor: true, propose: true,
  acceptProposal: true, rejectProposal: true, applyEdits: true, addUndoBoundary: true, undo: true, redo: true,
} satisfies Record<PptxWorkerEditorOperation['method'], true>;

export interface WorkerInputCoordinator extends PptxInputCoordinator {
  access(peer: PptxWorkerEditorAccess): PptxWorkerEditorAccess;
  flush(): Promise<void>;
  beginTyping(target: string, time?: number): void;
  endTyping(): void;
  breakTyping(): void;
}

export function createWorkerInputCoordinator(hooks: {
  session(): PptxWorkerEditorSession | null;
  gestureActive(): boolean;
  readOnly?(): boolean;
  changed?(): void;
}): WorkerInputCoordinator {
  let typing = false;
  let previous: { target: string; time: number } | undefined;
  const require = () => {
    const session = hooks.session();
    if (!session) throw new PptxPeerNotReadyError();
    if (session.state.stage === 'disposed') throw new PptxWorkerEditorDisposedError();
    if (session.failure) throw session.failure;
    if (!session.hydrated) throw new PptxPeerNotReadyError();
    return session;
  };
  const run = <T,>(operation: () => T | Promise<T>): Promise<T> => {
    try { require(); return Promise.resolve(operation()); }
    catch (error) { return Promise.reject(error); }
  };
  const writable = () => {
    const session = require();
    if (hooks.readOnly?.()) throw new PptxCommandAdmissionError('read-only');
    return session;
  };
  const boundary = () => writable().apply({ method: 'addUndoBoundary', args: [] });
  return {
    busy: () => false,
    keyboardQueued: () => false,
    input: (operation) => run(operation),
    run,
    reset() { typing = false; previous = undefined; },
    breakTyping() { previous = undefined; },
    beginTyping(target, time = Date.now()) {
      writable();
      if (!previous || previous.target !== target || time - previous.time > 500) boundary();
      previous = { target, time };
      typing = true;
    },
    endTyping() { typing = false; },
    flush() {
      return run(() => {
        if (hooks.gestureActive()) throw new Error('Finish the pointer gesture before flushing input');
        return require().flush();
      });
    },
    access(peer) {
      const owner = require();
      const requireAccess = () => {
        if (owner.state.stage === 'disposed' || hooks.session() !== owner) throw new PptxWorkerEditorDisposedError();
        return require();
      };
      const access = Object.create(null) as PptxWorkerEditorAccess;
      for (const key of Reflect.ownKeys(peer)) {
        Object.defineProperty(access, key, {
          enumerable: typeof key === 'string',
          get: () => {
            requireAccess();
            if (typeof key !== 'string' || !Object.hasOwn(mutations, key)) return Reflect.get(peer, key);
            return (...args: unknown[]) => {
              const session = requireAccess();
              if (hooks.readOnly?.()) {
                if (key === 'applyEdits') return {
                  ok: false, version: peer.version(),
                  failure: { code: 'read-only', message: 'The editor is read-only' },
                };
                throw new PptxCommandAdmissionError('read-only');
              }
              if (!typing && key !== 'addUndoBoundary') previous = undefined;
              const discrete = !typing && !['addUndoBoundary', 'undo', 'redo', 'applyEdits'].includes(key);
              if (discrete) { previous = undefined; boundary(); }
              try {
                const reply = session.apply({ method: key, args } as PptxWorkerEditorOperation);
                return reply.outcome.result;
              } finally {
                if (discrete && !session.failure) boundary();
                hooks.changed?.();
              }
            };
          },
        });
      }
      return Object.freeze(access);
    },
  };
}

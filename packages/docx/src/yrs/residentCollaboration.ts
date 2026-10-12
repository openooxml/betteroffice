import type { CollaborationTextInsertion } from '../collaboration/types';
import { MAX_COLLABORATION_FRAME_BYTES } from '../collaboration/protocol';
import type { ResidentEngineSession } from './residentEngineSession';
import type { ResidentCommandHooksFactory, DocxCommandIntent } from './residentCommandProtocol';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { DisplayPositionIndex } from './displayPositionIndex';
import { createDisplayListQueries, type DisplayListRect } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';

export interface ResidentPresencePeer {
  clientId: number;
  clock: number;
  cursor?: { story: string; anchor: Uint8Array; head: Uint8Array };
  inferred?: { story: string; paraId: string; offset: number };
}

export interface ResidentPresenceGeometry {
  version: string;
  layoutRevision: number;
  frameEpoch: number;
  peers: Array<{ clientId: number; clock: number; status: 'ok' | 'unresolved';
    caret: (DisplayListRect & { pageId: string }) | null;
    rects: Array<DisplayListRect & { pageId: string }> }>;
}

export interface ResidentCollaborationUpdates {
  updates: ArrayBuffer[];
  inferences: Array<{ seq: number; insertion: CollaborationTextInsertion }>;
}

const inferredAnchors = new WeakMap<ResidentEngineSession, Map<number, { key: string; sticky: { story: string; encoded: Uint8Array } }>>();
const inferenceKey = (loc: { story: string; paraId: string; offset: number }) => JSON.stringify([loc.story, loc.paraId, loc.offset]);

type Hooks = ReturnType<ResidentCommandHooksFactory>;

export function createResidentCollaboration(engine: ResidentEngineSession) {
  let disposed = false;
  const anchors = new Map<number, { key: string; sticky: { story: string; encoded: Uint8Array } }>();
  inferredAnchors.set(engine, anchors);
  const updates: ArrayBuffer[] = [];
  const inferences: ResidentCollaborationUpdates['inferences'] = [];
  const unsubscribe = engine.onUpdate((update, origin) => {
    if (!disposed && origin !== 'remote') updates.push(update.slice().buffer);
  });
  return {
    drain(): ResidentCollaborationUpdates | undefined {
      if (!updates.length && !inferences.length) return undefined;
      return { updates: updates.splice(0), inferences: inferences.splice(0) };
    },
    wrap(hooks: Hooks): Hooks {
      const state = () => ({ selection: hooks.selection(), editState: hooks.editState() });
      return {
        ...hooks,
        preflight(command, context) {
          if (command.kind !== 'applyRemoteUpdate') return hooks.preflight(command, context);
          const payload = command.payload as { update?: unknown } | null;
          if (disposed || !payload || !(payload.update instanceof ArrayBuffer) || !payload.update.byteLength ||
            payload.update.byteLength > MAX_COLLABORATION_FRAME_BYTES) {
            return { status: 'refused', reason: 'invalid-payload', ...state() };
          }
          return { status: 'ready', prepared: { command: command as DocxCommandIntent,
            selection: hooks.selection(), range: null } };
        },
        apply(prepared, context) {
          if (prepared.command.kind !== 'applyRemoteUpdate') return hooks.apply(prepared, context);
          const before = engine.proposalEngine.version();
          try {
            const insertion = engine.applyUpdate(new Uint8Array(prepared.command.payload.update));
            if (insertion) {
              inferences.push({ seq: prepared.command.seq, insertion });
              const loc = { story: insertion.story, paraId: insertion.paraId, offset: insertion.endOffset };
              try { anchors.set(insertion.clientId, { key: inferenceKey(loc), sticky: engine.encodeStickyPosition(loc) }); } catch {}
              if (anchors.size > 64) anchors.delete(anchors.keys().next().value!);
            }
            return { status: engine.proposalEngine.version() === before ? 'noop' : 'applied', ...state() };
          } catch (error) {
            if (error instanceof WebAssembly.RuntimeError || engine.proposalEngine.version() !== before) throw error;
            return { status: 'refused', reason: 'invalid-update', ...state() };
          }
        },
      };
    },
    dispose() { if (disposed) return; disposed = true; unsubscribe(); updates.length = 0; inferences.length = 0; anchors.clear(); inferredAnchors.delete(engine); },
  };
}

export function readResidentCollaborationState(engine: ResidentEngineSession, stateVector?: Uint8Array): Uint8Array {
  return engine.encodeStateAsUpdate(stateVector);
}

export async function readResidentPresence(
  engine: ResidentEngineSession, frame: RetainedFrame, peers: readonly ResidentPresencePeer[],
  pages: readonly [number, number], layoutRevision: number
): Promise<ResidentPresenceGeometry> {
  const native = await loadRustDisplayListQueryEngine();
  const queries = createDisplayListQueries(frame.displayList, native);
  const index = new DisplayPositionIndex({ ...engine.geometryReader, selectionText: engine.selectionText });
  const visible = (rect: DisplayListRect) => rect.pageIndex >= pages[0] && rect.pageIndex < pages[1];
  const attach = (rect: DisplayListRect) => ({ ...rect, pageId: frame.pages.find(page => page.pageIndex === rect.pageIndex)!.pageId.toString() });
  try {
    const found = peers.slice(0, 32).map(peer => {
      const unresolved = { clientId: peer.clientId, clock: peer.clock, status: 'unresolved' as const, caret: null, rects: [] };
      const cached = inferredAnchors.get(engine)?.get(peer.clientId);
      const inferred = peer.inferred && cached?.key === inferenceKey(peer.inferred) ? engine.resolveStickyPosition(cached.sticky) : peer.inferred;
      const anchor = peer.inferred ? inferred : (peer.cursor && engine.resolveStickyPosition({ story: peer.cursor.story, encoded: peer.cursor.anchor }));
      const head = peer.inferred ? inferred : (peer.cursor && engine.resolveStickyPosition({ story: peer.cursor.story, encoded: peer.cursor.head }));
      if (!anchor || !head || anchor.story !== head.story) return unresolved;
      const root = head.story === 'body' || head.story.startsWith('body:') ? 'body' : head.story.split(':').slice(0, 2).join(':');
      const from = index.positionOf(anchor, root), to = index.positionOf(head, root);
      if (from === null || to === null) return unresolved;
      const [kind, id] = root.split(':');
      const start = Math.min(from, to), end = Math.max(from, to);
      const rects = start === end ? [] : root === 'body' ? (queries.rangeRectsOnPages?.(start, end, pages[0], pages[1] - 1) ?? queries.rangeRects(start, end))
        : kind === 'hf' ? [...queries.hfRangeRects('header', id!, start, end), ...queries.hfRangeRects('footer', id!, start, end)]
        : kind === 'fn' || kind === 'en' ? queries.noteRangeRects(kind === 'fn' ? 'footnote' : 'endnote', Number(id), start, end) : [];
      const carets = root === 'body' ? [queries.caretRect(to)]
        : kind === 'hf' ? [...queries.hfCaretRects('header', id!, to), ...queries.hfCaretRects('footer', id!, to)]
        : kind === 'fn' || kind === 'en' ? queries.noteCaretRects(kind === 'fn' ? 'footnote' : 'endnote', Number(id), to) : [];
      const caret = carets.find(rect => rect && visible(rect));
      return { clientId: peer.clientId, clock: peer.clock, status: 'ok' as const,
        caret: caret ? attach(caret) : null, rects: rects.filter(visible).slice(0, 64).map(attach) };
    });
    const state = queries.sourceState();
    if (state.status === 'error') throw new Error(`Presence layout is unavailable: ${String(state.error)}`);
    return { version: engine.proposalEngine.version(), layoutRevision, frameEpoch: frame.frameEpoch, peers: found };
  } finally { queries.dispose(); }
}

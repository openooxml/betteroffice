import { CollaborationError, type CollaborationTextInsertion } from '../collaboration/types';
import { stateVectorClock } from '../collaboration/protocol';
import { CollaborationProvider, providerArrival, providerLeave } from '../collaboration/provider';
import { createCollaborationPipeHost, type PipeMergeResult } from '../../../../shared/collaboration-pipe/worker';
import type { PipeOutputBatch } from '../../../../shared/collaboration-pipe/protocol';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandHooksFactory } from './residentCommandProtocol';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { DisplayPositionIndex } from './displayPositionIndex';
import { createDisplayListQueries, type DisplayListRect } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';

export interface DocxPresencePeer {
  clientId: number;
  clock: number;
  cursor?: { story: string; anchor: Uint8Array; head: Uint8Array };
  inferred?: { story: string; paraId: string; offset: number };
}

export interface DocxPresenceGeometry {
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
  /** This session's own state-vector clock after the drained edits and merges. */
  clock: number;
  /** Output of the providers this worker runs, for the threads that own their transports. */
  pipe?: PipeOutputBatch[];
}

const inferredAnchors = new WeakMap<ResidentEngineSession, Map<number, { key: string; sticky: { story: string; encoded: Uint8Array } }>>();
const inferenceKey = (loc: { story: string; paraId: string; offset: number }) => JSON.stringify([loc.story, loc.paraId, loc.offset]);

type Hooks = ReturnType<DocxCommandHooksFactory>;

export function createResidentCollaboration(engine: ResidentEngineSession, options: { onPipeOutput?: () => void } = {}) {
  let disposed = false;
  const anchors = new Map<number, { key: string; sticky: { story: string; encoded: Uint8Array } }>();
  inferredAnchors.set(engine, anchors);
  const updates: ArrayBuffer[] = [];
  const inferences: ResidentCollaborationUpdates['inferences'] = [];
  let merged = false;
  const unsubscribe = engine.onUpdate((update, origin) => {
    if (!disposed && origin !== 'remote') updates.push(update.slice().buffer);
  });
  const remember = (insertion: CollaborationTextInsertion) => {
    const loc = { story: insertion.story, paraId: insertion.paraId, offset: insertion.endOffset };
    try { anchors.set(insertion.clientId, { key: inferenceKey(loc), sticky: engine.encodeStickyPosition(loc) }); } catch {}
    if (anchors.size > 64) anchors.delete(anchors.keys().next().value!);
  };
  const pipe = createCollaborationPipeHost({
    clientId: engine.clientId,
    read: {
      encodeStateVector: () => engine.encodeStateVector(),
      encodeStateAsUpdate: (vector) => engine.encodeStateAsUpdate(vector),
      encodeStateAsUpdateChunks: (vector, maxChunkBytes) => engine.encodeStateAsUpdateChunks(vector, maxChunkBytes),
    },
    onUpdate: (listener) => engine.onUpdate(listener),
    merge(peerUpdates, single): PipeMergeResult {
      const before = engine.proposalEngine.version();
      try {
        const insertion = single ? engine.applyUpdate(peerUpdates[0]!) : (engine.applyUpdateChunks(peerUpdates), null);
        merged = true;
        if (insertion) remember(insertion);
        return { applied: insertion };
      } catch (error) {
        if (error instanceof WebAssembly.RuntimeError || engine.proposalEngine.version() !== before) throw error;
        return { failed: error };
      }
    },
    createProvider: (replica, transport, providerOptions) => new CollaborationProvider(replica, transport, providerOptions),
    arrive: providerArrival,
    leave: providerLeave,
    transportError: (error) => new CollaborationError(error.code, error.message),
    onOutput: options.onPipeOutput,
  });
  return {
    drain(): (ResidentCollaborationUpdates & { transfer: ArrayBuffer[] }) | undefined {
      const piped = pipe.drain();
      if (!updates.length && !inferences.length && !merged && !piped) return undefined;
      merged = false;
      return { clock: stateVectorClock(engine.encodeStateVector(), engine.clientId), updates: updates.splice(0), inferences: inferences.splice(0),
        ...(piped && { pipe: piped.batches }), transfer: piped?.transfer ?? [] };
    },
    wrap(hooks: Hooks): Hooks {
      const state = () => ({ selection: hooks.selection(), editState: hooks.editState() });
      return {
        ...hooks,
        apply(prepared, context) {
          const { command } = prepared;
          if (command.kind !== 'applyRemoteUpdate' && command.kind !== 'applyRemoteUpdateChunks' && command.kind !== 'collabPipe') {
            return hooks.apply(prepared, context);
          }
          if (disposed) return { status: 'refused', reason: 'unsupported-command', ...state() };
          const before = engine.proposalEngine.version();
          const since = engine.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
          if (command.kind === 'collabPipe') {
            pipe.run(command.payload);
            if (engine.proposalEngine.version() === before) return { status: 'noop', ...state() };
            engine.markProjectionStories(engine.storiesChangedSince(since).stories);
            hooks.afterRemoteMerge?.(context);
            return { status: 'applied', ...state() };
          }
          try {
            if (command.kind === 'applyRemoteUpdateChunks') engine.applyUpdateChunks(command.payload.updates.map(update => new Uint8Array(update)));
            const insertion = command.kind === 'applyRemoteUpdate' ? engine.applyUpdate(new Uint8Array(command.payload.update)) : null;
            merged = true;
            if (engine.proposalEngine.version() !== before) engine.markProjectionStories(engine.storiesChangedSince(since).stories);
            if (insertion) {
              inferences.push({ seq: command.seq, insertion });
              remember(insertion);
            }
            if (engine.proposalEngine.version() === before) return { status: 'noop', ...state() };
            hooks.afterRemoteMerge?.(context);
            return { status: 'applied', ...state() };
          } catch (error) {
            if (error instanceof WebAssembly.RuntimeError || engine.proposalEngine.version() !== before) throw error;
            return { status: 'refused', reason: 'invalid-update', ...state() };
          }
        },
      };
    },
    dispose() { if (disposed) return; disposed = true; pipe.dispose(); unsubscribe(); updates.length = 0; inferences.length = 0; anchors.clear(); inferredAnchors.delete(engine); },
  };
}

export function readResidentCollaborationState(engine: ResidentEngineSession, stateVector?: Uint8Array): Uint8Array {
  return engine.encodeStateAsUpdate(stateVector);
}

export async function readResidentPresence(
  engine: ResidentEngineSession, frame: RetainedFrame, peers: readonly DocxPresencePeer[],
  pages: readonly [number, number], layoutRevision: number
): Promise<DocxPresenceGeometry> {
  const native = await loadRustDisplayListQueryEngine();
  const queries = createDisplayListQueries(frame.displayList, native);
  const index = new DisplayPositionIndex({ ...engine.geometryReader, selectionText: engine.selectionText });
  const visible = (rect: DisplayListRect) => rect.pageIndex >= pages[0] && rect.pageIndex < pages[1];
  const attach = (rect: DisplayListRect) => {
    const page = frame.pages.find(page => page.pageIndex === rect.pageIndex);
    return page ? [{ ...rect, pageId: page.pageId.toString() }] : [];
  };
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
      const caret = carets.flatMap(rect => rect && visible(rect) ? attach(rect) : [])[0] ?? null;
      return { clientId: peer.clientId, clock: peer.clock, status: 'ok' as const,
        caret, rects: rects.filter(visible).flatMap(attach).slice(0, 64) };
    });
    const state = queries.sourceState();
    if (state.status === 'error') throw new Error(`Presence layout is unavailable: ${String(state.error)}`);
    return { version: engine.proposalEngine.version(), layoutRevision, frameEpoch: frame.frameEpoch, peers: found };
  } finally { queries.dispose(); }
}

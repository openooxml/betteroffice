import type { EditSessionState } from '../../../../../../shared/office-session/commands/types';
import type { DisplayPage } from '@betteroffice/docx/layout/render';
import type {
  DocxCommandFrame,
  DocxCommandPageIdentity,
  DocxCommandSelection,
  DocxEditState,
  DocxEchoRect,
  DocxEchoTransform,
  DocxFrameReservations,
} from '@betteroffice/docx/yrs';

/** @internal */
export interface PeerlessViewIdentity {
  session: string;
  incarnation: number;
  epoch: number;
}

/** @internal Worker-minted bytes retained without decoding on main. */
export interface PeerlessSelectionAnchors {
  anchor: Uint8Array;
  focus: Uint8Array;
}

/** @internal Installed geometry is independent of acknowledged worker selection. */
export interface PeerlessFrameGeometry {
  documentVersion?: string;
  identity: PeerlessViewIdentity;
  coversThroughSeq: number;
  frameEpoch: number;
  layoutRevision: number;
  /** The worker pagination run the pages show; frames that only move the caret keep it. */
  layoutEpoch?: number;
  documentWidth?: number;
  layoutProvisional?: boolean;
  /** The reservation set the layout includes. */
  reservations?: DocxFrameReservations;
  pages: readonly {
    page: DocxCommandPageIdentity;
    bounds: DocxEchoRect;
    transform: DocxEchoTransform;
    displayPage?: DisplayPage;
  }[];
  caret: DocxEchoRect | null;
  selectionFocus?: DocxEchoRect | null;
  selectionRects: readonly DocxEchoRect[];
  parts?: DocxCommandFrame['payload']['parts'];
  objects?: DocxCommandFrame['payload']['geometry'];
}

/** @internal */
export interface PeerlessProvisionalCaret {
  readonly throughSeq: number | undefined;
  pageIndex: number;
  rect: DocxEchoRect;
}

/** @internal */
export interface PeerlessViewSnapshot {
  identity: PeerlessViewIdentity;
  version: number;
  throughSeq: number;
  presentedThroughSeq: number;
  workerSelection: DocxCommandSelection | null | undefined;
  anchors: PeerlessSelectionAnchors | null;
  editState: DocxEditState | undefined;
  geometry: PeerlessFrameGeometry | null;
  provisionalCaret?: PeerlessProvisionalCaret | null;
}

/** @internal Stable snapshots and subscriptions for useSyncExternalStore. */
export interface PeerlessViewStore {
  subscribe(listener: () => void): () => void;
  getSnapshot(): PeerlessViewSnapshot;
  getServerSnapshot(): PeerlessViewSnapshot;
  acceptWorkerState(
    identity: PeerlessViewIdentity,
    state: EditSessionState<DocxCommandSelection | null, DocxEditState>
  ): void;
  installFrameGeometry(geometry: PeerlessFrameGeometry): void;
  setProvisionalCaret(caret: PeerlessProvisionalCaret | null): void;
  dispose(): void;
}

/** @internal */
export function createPeerlessViewStore(initial: PeerlessViewSnapshot): PeerlessViewStore {
  const anchors = (selection: PeerlessViewSnapshot['workerSelection']): PeerlessSelectionAnchors | null =>
    selection?.anchor.sticky && selection.focus.sticky
      ? { anchor: selection.anchor.sticky.slice(), focus: selection.focus.sticky.slice() } : null;
  let snapshot: PeerlessViewSnapshot = {
    ...initial, provisionalCaret: null, identity: { ...initial.identity }, anchors: anchors(initial.workerSelection),
    workerSelection: initial.workerSelection === undefined ? undefined : structuredClone(initial.workerSelection),
  };
  let selectionInput = initial.workerSelection;
  let provisionalRevision = {
    version: initial.version, documentVersion: initial.geometry?.documentVersion, layoutRevision: initial.geometry?.layoutRevision,
  };
  let disposed = false;
  const listeners = new Set<() => void>();
  const sameIdentity = (identity: PeerlessViewIdentity) =>
    identity.session === snapshot.identity.session && identity.incarnation === snapshot.identity.incarnation &&
    identity.epoch === snapshot.identity.epoch;
  const publish = (next: PeerlessViewSnapshot) => {
    snapshot = next;
    for (const listener of listeners) listener();
  };
  return {
    subscribe(listener) {
      if (disposed) return () => {};
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    getSnapshot: () => snapshot,
    getServerSnapshot: () => snapshot,
    acceptWorkerState(identity, state) {
      if (disposed || state.epoch !== identity.epoch || identity.session !== snapshot.identity.session ||
        identity.incarnation < snapshot.identity.incarnation || identity.epoch < snapshot.identity.epoch) return;
      const replacement = !sameIdentity(identity);
      if (!replacement && (state.version < snapshot.version || state.throughSeq < snapshot.throughSeq)) return;
      const selected = state.selection === undefined && !replacement ? snapshot.workerSelection : state.selection;
      const editState = state.editState === undefined && !replacement ? snapshot.editState : state.editState;
      if (!replacement && state.version === snapshot.version && state.throughSeq === snapshot.throughSeq &&
        (state.selection === undefined || state.selection === selectionInput) && editState === snapshot.editState) return;
      selectionInput = selected;
      publish({
        identity: { ...identity }, version: state.version, throughSeq: state.throughSeq,
        presentedThroughSeq: replacement ? -1 : snapshot.presentedThroughSeq,
        workerSelection: selected === undefined ? undefined : structuredClone(selected),
        anchors: anchors(selected), editState,
        geometry: replacement ? null : snapshot.geometry,
        provisionalCaret: replacement || state.version > provisionalRevision.version ? null : snapshot.provisionalCaret,
      });
    },
    installFrameGeometry(geometry) {
      if (disposed || !sameIdentity(geometry.identity) || geometry.coversThroughSeq < snapshot.presentedThroughSeq ||
        (snapshot.geometry && (geometry.frameEpoch <= snapshot.geometry.frameEpoch ||
          geometry.layoutRevision < snapshot.geometry.layoutRevision))) return;
      const seq = snapshot.provisionalCaret?.throughSeq;
      publish({ ...snapshot, presentedThroughSeq: geometry.coversThroughSeq, geometry,
        provisionalCaret: (seq !== undefined && geometry.coversThroughSeq >= seq) ||
          geometry.documentVersion !== provisionalRevision.documentVersion ||
          (provisionalRevision.layoutRevision !== undefined && geometry.layoutRevision > provisionalRevision.layoutRevision)
          ? null : snapshot.provisionalCaret });
    },
    setProvisionalCaret(caret) {
      if (disposed || snapshot.provisionalCaret === caret ||
        (caret?.throughSeq !== undefined && snapshot.geometry && snapshot.geometry.coversThroughSeq >= caret.throughSeq)) return;
      provisionalRevision = {
        version: snapshot.version, documentVersion: snapshot.geometry?.documentVersion, layoutRevision: snapshot.geometry?.layoutRevision,
      };
      publish({ ...snapshot, provisionalCaret: caret });
    },
    dispose() {
      disposed = true;
      listeners.clear();
    },
  };
}

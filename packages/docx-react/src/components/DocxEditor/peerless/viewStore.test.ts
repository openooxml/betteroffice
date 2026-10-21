import { expect, test } from 'bun:test';
import { createPeerlessViewStore, type PeerlessFrameGeometry } from './viewStore';

test('d2a_ack_selection_waits_for_matching_frame_geometry', () => {
  const identity = { session: 'view', incarnation: 1, epoch: 1 };
  const selection = (offset: number) => ({
    anchor: { story: 'body', paraId: 'p', offset, sticky: new Uint8Array([offset, 1]) },
    focus: { story: 'body', paraId: 'p', offset, sticky: new Uint8Array([offset, 2]) },
  });
  const geometry = (coversThroughSeq: number, frameEpoch: number): PeerlessFrameGeometry => ({
    identity, coversThroughSeq, frameEpoch, layoutRevision: frameEpoch,
    pages: [], caret: { x: frameEpoch, y: 0, width: 1, height: 12 }, selectionRects: [],
  });
  const old = geometry(1, 1);
  const store = createPeerlessViewStore({
    identity, version: 1, throughSeq: 1, presentedThroughSeq: 1,
    workerSelection: selection(1), anchors: null, editState: undefined, geometry: old,
  });
  let notifications = 0;
  store.subscribe(() => notifications++);
  const confirmed = selection(2);
  store.acceptWorkerState(identity, { epoch: 1, version: 2, throughSeq: 2, presentedThroughSeq: 1, selection: confirmed, editState: undefined });
  expect(store.getSnapshot()).toMatchObject({ throughSeq: 2, presentedThroughSeq: 1, workerSelection: confirmed, geometry: old });
  expect(store.getSnapshot().anchors).toEqual({ anchor: confirmed.anchor.sticky, focus: confirmed.focus.sticky });
  const matching = geometry(2, 2);
  store.installFrameGeometry(matching);
  expect(store.getSnapshot().geometry).toEqual(matching);
  expect(store.getSnapshot().presentedThroughSeq).toBe(2);
  const stable = store.getSnapshot();
  store.acceptWorkerState(identity, { epoch: 1, version: 1, throughSeq: 1, presentedThroughSeq: 1, selection: selection(0), editState: undefined });
  store.installFrameGeometry(old);
  store.installFrameGeometry({ ...geometry(3, 3), identity: { ...identity, incarnation: 0 } });
  expect(store.getSnapshot()).toBe(stable);
  store.installFrameGeometry(geometry(2, 3));
  expect(store.getSnapshot().geometry?.frameEpoch).toBe(3);
  expect(store.getSnapshot().workerSelection).toEqual(confirmed);
  store.acceptWorkerState(identity, { epoch: 1, version: 2, throughSeq: 3, presentedThroughSeq: 2, selection: undefined, editState: undefined });
  expect(store.getSnapshot().workerSelection).toEqual(confirmed);
  expect(store.getServerSnapshot()).toBe(store.getSnapshot());
  store.dispose();
  const count = notifications;
  store.installFrameGeometry(geometry(3, 4));
  expect(notifications).toBe(count);
});

test('provisional caret follows its ticket sequence and retires at covering geometry', () => {
  const identity = { session: 'view', incarnation: 1, epoch: 1 };
  const geometry: PeerlessFrameGeometry = {
    identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1,
    pages: [], caret: null, selectionRects: [],
  };
  const store = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0,
    workerSelection: null, anchors: null, editState: undefined, geometry });
  let seq: number | undefined;
  const caret = { get throughSeq() { return seq; }, pageIndex: 7, rect: { x: 30, y: 1200, width: 1, height: 18 } };
  store.setProvisionalCaret(caret);
  expect(store.getSnapshot().provisionalCaret).toBe(caret);
  store.installFrameGeometry({ ...geometry, frameEpoch: 2 });
  expect(store.getSnapshot().provisionalCaret).toBe(caret);
  seq = 1;
  store.acceptWorkerState(identity, { epoch: 1, version: 0, throughSeq: 1, presentedThroughSeq: 0,
    selection: undefined, editState: undefined });
  expect(store.getSnapshot().provisionalCaret?.throughSeq).toBe(1);
  store.installFrameGeometry({ ...geometry, frameEpoch: 3, coversThroughSeq: 1 });
  expect(store.getSnapshot().provisionalCaret).toBeNull();
  store.setProvisionalCaret(caret);
  expect(store.getSnapshot().provisionalCaret).toBeNull();
  seq = 2;
  store.setProvisionalCaret(caret);
  store.acceptWorkerState({ ...identity, incarnation: 2 }, { epoch: 1, version: 0, throughSeq: 0,
    presentedThroughSeq: -1, selection: null, editState: undefined });
  expect(store.getSnapshot().provisionalCaret).toBeNull();
  store.dispose();
});

for (const change of ['document version', 'layout revision', 'frame document version']) {
  test(`a newer ${change} retires a provisional caret without covering its sequence`, () => {
    const identity = { session: 'view', incarnation: 1, epoch: 1 };
    const geometry: PeerlessFrameGeometry = {
      identity, documentVersion: 'doc-1', coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1,
      pages: [], caret: null, selectionRects: [],
    };
    const store = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0,
      workerSelection: null, anchors: null, editState: undefined, geometry });
    const caret = { throughSeq: 2, pageIndex: 0, rect: { x: 30, y: 0, width: 1, height: 18 } };
    store.setProvisionalCaret(caret);
    if (change === 'document version') {
      store.acceptWorkerState(identity, { epoch: 1, version: 1, throughSeq: 0, presentedThroughSeq: 0,
        selection: undefined, editState: undefined });
    } else {
      store.installFrameGeometry({ ...geometry, frameEpoch: 2,
        ...(change === 'layout revision' ? { layoutRevision: 2 } : { documentVersion: 'doc-2' }) });
    }
    expect(store.getSnapshot().provisionalCaret).toBeNull();
    expect(store.getSnapshot().presentedThroughSeq).toBe(0);
    store.setProvisionalCaret(caret);
    store.installFrameGeometry({ ...store.getSnapshot().geometry!, frameEpoch: 3 });
    expect(store.getSnapshot().provisionalCaret).toBe(caret);
    store.dispose();
  });
}

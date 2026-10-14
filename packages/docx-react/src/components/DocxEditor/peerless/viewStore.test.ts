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

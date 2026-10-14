import { expect, mock, test } from 'bun:test';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createSharedPresentationCommandClient } from './commandClient';
import { createPresentationCommandProjection } from './commandSharedAdapter';
import type { PresentationCommandOpenResult } from './commandProtocol';

test('rejected shared frames close each embedded bitmap once', async () => {
  class Bitmap {
    close = mock(() => {});
  }
  const original = Object.getOwnPropertyDescriptor(globalThis, 'ImageBitmap');
  Object.defineProperty(globalThis, 'ImageBitmap', { configurable: true, value: Bitmap });
  const first = new Bitmap();
  const second = new Bitmap();
  const state: PresentationCommandOpenResult = {
    session: 'session', incarnation: 'worker', epoch: 1, throughSeq: 0, version: 'v0', selection: null,
    editState: { readOnly: false, canUndo: false, canRedo: false, toolbar: { target: null,
      storedMarks: {}, formatting: null, alignment: null, hyperlink: null, table: null, canEditText: false } },
    outcomes: [], dirtySlideIds: [], activeSlideId: 'slide', slides: [], size: { width: 960, height: 540 },
    frame: { slideId: 'slide', slideIndex: 0, slideRevision: 0, coversThroughSeq: 0, documentEpoch: 1,
      incarnation: 'worker', epoch: 1, version: 'v0', displayList: new ArrayBuffer(0), media: [] },
  };
  const projection = createPresentationCommandProjection();
  projection.initialize(state.version);
  let deliver: (message: unknown) => void = () => {};
  const transport: SessionTransport = {
    post(message) {
      if (!isClientMessage(message) || message.kind !== 'call') return;
      const value = message.method === 'openCommandEditor' ? { ...state, shared: projection.ack(state) }
        : { ...projection.ack({ ...state, throughSeq: 1,
          outcomes: [{ seq: 1, status: 'noop', committed: false }],
          frame: Object.assign({ ...state.frame!, coversThroughSeq: 1, epoch: 2 }, {
            bitmaps: [first, first, second],
          }) }), incarnation: 99 };
      deliver({ protocol: 1, kind: 'reply', id: message.id, ok: true, value });
    },
    listen(listener) { deliver = listener; return () => { deliver = () => {}; }; },
    onError() { return () => {}; },
    close() {},
  };
  const client = await createSharedPresentationCommandClient(new Uint8Array(),
    { session: state.session, incarnation: state.incarnation }, transport);
  try {
    void client.dispatch({ kind: 'setSelection', payload: { inputTime: 1, undoGroup: 'selection', selection: null } })
      .catch(() => {});
    for (let turn = 0; turn < 12; turn++) await Promise.resolve();
    expect(first.close).toHaveBeenCalledTimes(1);
    expect(second.close).toHaveBeenCalledTimes(1);
  } finally {
    await client.dispose();
    if (original) Object.defineProperty(globalThis, 'ImageBitmap', original);
    else Reflect.deleteProperty(globalThis, 'ImageBitmap');
  }
});

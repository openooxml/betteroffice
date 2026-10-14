import { expect, mock, test } from 'bun:test';
import * as session from '../../../../shared/office-session/commands/session';
import { isClientMessage, type ClientMessage } from '../../../../shared/office-session/protocol';
import { detachedBuffers, freezingEdits, type FrozenEnvelope } from '../../../../shared/office-session/testing/frozen';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createPresentationCommandProjection } from './commandSharedAdapter';
import type { PresentationCommandFailure, PresentationCommandOpenResult } from './commandProtocol';

const actual = { ...session };
const received: FrozenEnvelope[] = [];
mock.module('../../../../shared/office-session/commands/session', () => ({
  ...actual, createWorkerEditSession: freezingEdits(actual.createWorkerEditSession, received),
}));
const { createSharedPresentationCommandClient } = await import('./commandClient');

const state: PresentationCommandOpenResult = {
  session: 'session', incarnation: 'worker', epoch: 1, throughSeq: 0, version: 'v0', selection: null,
  editState: { readOnly: false, canUndo: false, canRedo: false, toolbar: { target: null,
    storedMarks: {}, formatting: null, alignment: null, hyperlink: null, table: null, canEditText: false } },
  outcomes: [], dirtySlideIds: [], activeSlideId: 'slide', slides: [], size: { width: 960, height: 540 },
  frame: { slideId: 'slide', slideIndex: 0, slideRevision: 0, coversThroughSeq: 0, documentEpoch: 1,
    incarnation: 'worker', epoch: 1, version: 'v0', displayList: new ArrayBuffer(0), media: [] },
};

function copyingWorker() {
  const projection = createPresentationCommandProjection();
  projection.initialize(state.version);
  const edits: { message: Extract<ClientMessage, { kind: 'call' }>; transfer: Transferable[] }[] = [];
  let deliver: (message: unknown) => void = () => {};
  const transport: SessionTransport = {
    post(message, transfer = []) {
      const delivered = structuredClone(message, { transfer });
      if (!isClientMessage(delivered) || delivered.kind !== 'call') return;
      let value: unknown;
      if (delivered.method === 'openCommandEditor') value = { ...state, shared: projection.ack(state) };
      else if (delivered.method === 'edit') {
        edits.push({ message: delivered, transfer });
        const seq = (delivered.args[0] as { seq: number }).seq;
        projection.commit(`v${seq}`);
        value = projection.ack({ ...state, throughSeq: seq, version: `v${seq}`, frame: null,
          outcomes: [{ seq, status: 'applied', committed: true }] });
      } else return;
      queueMicrotask(() => { deliver({ protocol: 1, kind: 'reply', id: delivered.id, ok: true, value }); });
    },
    listen(listener) { deliver = listener; return () => { deliver = () => {}; }; },
    onError() { return () => {}; },
    close() {},
  };
  return { transport, edits };
}

test('shared presentation edits leave frozen envelopes and their buffers untouched', async () => {
  const worker = copyingWorker();
  const client = await createSharedPresentationCommandClient(new Uint8Array(),
    { session: state.session, incarnation: state.incarnation }, worker.transport);
  const failures: PresentationCommandFailure[] = [];
  client.subscribeFailure((failure) => { failures.push(failure); });
  try {
    const picture = await client.dispatch({ kind: 'insertPicture', payload: { inputTime: 1, undoGroup: 'insert',
      slideId: 'slide', name: 'picture', rect: { x: 0, y: 0, width: 10, height: 10 }, contentType: 'image/png',
      bytes: Uint8Array.of(137, 80, 78, 71).buffer } });
    expect(picture.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    const font = await client.dispatch({ kind: 'registerFont', payload: { inputTime: 2, undoGroup: 'font',
      font: { family: 'Probe', bytes: Uint8Array.of(0, 1, 0, 0).buffer } } });
    expect(font.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    expect(received.map(({ envelope }) => [envelope.seq, envelope.kind, Object.isFrozen(envelope)])).toEqual([
      [1, 'insertPicture', true], [2, 'registerFont', true],
    ]);
    for (const { buffers } of received) {
      expect(buffers).toHaveLength(1);
      expect(detachedBuffers(buffers)).toEqual([]);
    }
    expect(worker.edits.map(({ transfer }) => transfer.length)).toEqual([1, 0]);
    expect(worker.edits.map(({ message }) => message.args[0])).toMatchObject([
      { seq: 1, kind: 'insertPicture', payload: { bytes: Uint8Array.of(137, 80, 78, 71).buffer } },
      { seq: 2, kind: 'registerFont', payload: { font: { bytes: Uint8Array.of(0, 1, 0, 0).buffer } } },
    ]);
    expect(failures).toEqual([]);
  } finally { await client.dispose(); }
});

import { expect, mock, test } from 'bun:test';
import { createPeerlessShellController } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

const editState = { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null, canUndo: false, canRedo: false, hyperlink: null, tableContext: null };

function fixture(readOnly = false, worker: { staleReview?: boolean; typing?: boolean } = {}) {
  const range = { story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 7 } };
  const comments: { id: string; author: string; date: string; body: unknown; done: boolean; parentId: null }[] = [];
  const projected = [{ id: 7, author: 'A', content: [] }];
  const viewStore = createPeerlessViewStore({ identity: { session: 's', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null, editState, geometry: null });
  const type = () => {
    const view = viewStore.getSnapshot();
    viewStore.acceptWorkerState(view.identity, { epoch: 1, version: view.version + 1, throughSeq: view.throughSeq + 1, presentedThroughSeq: 0, selection: null, editState });
  };
  const read = mock(async (request: { kind: string }) => {
    if (request.kind === 'readParagraphs' && worker.typing && comments.length) type();
    return { status: 'ok', version: 'native', value:
      request.kind === 'readParagraphs' ? { ok: true, version: 'native', paragraphs: [{ paraId: 'p', text: 'Seed' }] } :
      request.kind === 'commentInsertionRange' ? range : worker.staleReview ? null :
      { comments: structuredClone(comments), uiComments: projected, commentIds: { 'client:1': 7 }, frameEpoch: 1, layoutRevision: 1 } };
  });
  const submit = mock((intent: { payload: { author: string; date: string; body: unknown } }) => {
    comments.push({ id: 'client:1', ...intent.payload, done: false, parentId: null });
    const view = viewStore.getSnapshot();
    viewStore.acceptWorkerState(view.identity, { epoch: 1, version: view.version + 1, throughSeq: 1, presentedThroughSeq: 0, selection: null,
      editState: { ...editState, comment: { seq: 1, id: 7, comments: projected } } });
    const ack = Promise.resolve({ status: 'applied', seq: 1 }).then(outcome => { if (worker.typing) type(); return outcome; });
    return { status: 'submitted', ticket: { seq: 1, ack, presented: Promise.resolve('presented') } };
  });
  const changed = mock((_comments: unknown) => {});
  const controller = createPeerlessShellController({ read: read as never, readAt: read as never, focus() {}, zoom: 1,
    readOnly, onCommentsChange: changed });
  const binding = { session: { flush: async () => {} }, admission: {}, commands: { submit }, viewStore } as unknown as PeerlessEditSessionBinding;
  controller.attach(binding);
  return { controller, submit, range, changed, projected };
}

test('async comment creation uses native ranges and takes the assigned comment ID from its acknowledgement', async () => {
  const { controller, submit, range, changed, projected } = fixture();
  expect(await controller.ref.insertComment({ paraId: 'p', text: 'New note', author: 'A' })).toBe(7);
  expect(submit.mock.calls[0]![0]).toMatchObject({ kind: 'addComment', payload: { ranges: [range], author: 'A', expectVersion: 'native' } });
  expect(changed.mock.calls).toEqual([[projected]]);
});

test('async comment creation refuses read-only documents before admission', async () => {
  const { controller, submit } = fixture(true);
  expect(await controller.ref.insertComment({ paraId: 'p', text: 'New note', author: 'A' })).toBeNull();
  expect(submit).not.toHaveBeenCalled();
});

test('comment creation does not wait for a review layout the last save left behind', async () => {
  const { controller, submit, changed, projected } = fixture(false, { staleReview: true });
  expect(await controller.ref.insertComment({ paraId: 'p', text: 'New note', author: 'A' })).toBe(7);
  expect(submit).toHaveBeenCalledTimes(1);
  expect(changed.mock.calls).toEqual([[projected]]);
});

test('an applied comment reports its ID once when typing supersedes the next read', async () => {
  const { controller, submit, changed, projected } = fixture(false, { typing: true });
  expect(await controller.ref.insertComment({ paraId: 'p', text: 'New note', author: 'A' })).toBe(7);
  expect(submit).toHaveBeenCalledTimes(1);
  expect(changed.mock.calls).toEqual([[projected]]);
});

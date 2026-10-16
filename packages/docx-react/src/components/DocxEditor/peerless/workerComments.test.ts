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

type TestComment = { id: number; author: string; content: never[]; parentId?: number; done?: boolean };

function reviewFixture(options: { readOnly?: boolean; mode?: 'viewing' } = {}) {
  const ranges = [{ story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 7 } }];
  const host: TestComment[] = [{ id: 7, author: 'A', content: [] }];
  let comments: TestComment[] = structuredClone(host);
  let seq = 0;
  const viewStore = createPeerlessViewStore({ identity: { session: 's', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null, editState, geometry: null });
  const read = mock(async (request: { kind: string; target?: { kind: string; id?: number } }) => ({ status: 'ok', version: 'native', value:
    request.kind === 'readParagraphs' ? { ok: true, version: 'native', paragraphs: [] } :
    request.kind === 'reviewTarget' ? request.target?.id === 7 ? { id: 'native-7', ranges } : null :
    request.kind === 'reviewSnapshot' ? { comments: [], uiComments: structuredClone(comments), commentIds: { 'native-7': 7 } } : null }));
  const submit = mock((intent: { kind: string; payload: Record<string, unknown> & { author?: string; patch?: { done: boolean } } }) => {
    seq += 1;
    if (intent.kind === 'addComment') comments = [...comments, { id: 8, author: intent.payload.author!, content: [], parentId: 7 }];
    if (intent.kind === 'updateComment') comments = comments.map(comment => comment.id === 7 ? { ...comment, done: intent.payload.patch!.done } : comment);
    const view = viewStore.getSnapshot();
    const result = { seq, comments: structuredClone(comments) };
    viewStore.acceptWorkerState(view.identity, { epoch: 1, version: view.version + 1, throughSeq: seq, presentedThroughSeq: 0, selection: null,
      editState: { ...editState, ...(intent.kind === 'addComment' ? { comment: { ...result, id: 8 } } : { commentChange: result }) } });
    return { status: 'submitted', ticket: { seq, ack: Promise.resolve({ status: 'applied', seq }), presented: Promise.resolve('presented') } };
  });
  const changes: TestComment[][] = [];
  let changed: (() => void) | undefined;
  const controller = createPeerlessShellController({ read: read as never, readAt: read as never, focus() {}, zoom: 1,
    readOnly: options.readOnly, mode: options.mode, comments: () => host as never,
    onCommentsChange: (next) => { changes.push(next as TestComment[]); changed?.(); } });
  const binding = { session: { flush: async () => {} }, admission: {}, commands: { submit }, viewStore } as unknown as PeerlessEditSessionBinding;
  controller.attach(binding);
  return { controller, submit, ranges, host, changes, nextChange: () => new Promise<void>(resolve => { changed = resolve; }) };
}

test('a reply anchors where its parent is, resolves the new comment id and getComments shows it under its parent', async () => {
  const { controller, submit, ranges, host, changes } = reviewFixture();
  expect(controller.ref.getComments()).toEqual(host);
  expect(await controller.ref.insertCommentReply(7, 'Reply', 'B')).toBe(8);
  expect(submit.mock.calls[0]![0]).toMatchObject({ kind: 'addComment', payload: { ranges, author: 'B', body: 'Reply', parentId: 'native-7', expectVersion: 'native' } });
  expect(controller.ref.getComments()).toEqual([{ id: 7, author: 'A', content: [] }, { id: 8, author: 'B', content: [], parentId: 7 }]);
  expect(changes).toEqual([controller.ref.getComments() as never]);
});

test('resolveComment marks a comment done through the worker queue and a later change can reopen it', async () => {
  const { controller, submit, changes, nextChange } = reviewFixture();
  const resolved = nextChange();
  expect(controller.ref.resolveComment(7)).toBeUndefined();
  await resolved;
  expect(submit.mock.calls[0]![0]).toEqual({ kind: 'updateComment', payload: { id: 'native-7', patch: { done: true }, expectVersion: 'native' } });
  expect(controller.ref.getComments()).toEqual([{ id: 7, author: 'A', content: [], done: true }]);
  expect(await controller.setCommentDone(7, false)).toEqual({ ok: true, status: 'executed' });
  expect(controller.ref.getComments()).toEqual([{ id: 7, author: 'A', content: [], done: false }]);
  expect(changes).toHaveLength(2);
});

test('comment changes for an unknown comment refuse as values without admission', async () => {
  const { controller, submit, changes } = reviewFixture();
  expect(await controller.ref.insertCommentReply(99, 'Reply', 'B')).toBeNull();
  expect(await controller.setCommentDone(99, true)).toEqual({ ok: false, failure: { code: 'command-failed', message: 'unknown-comment' } });
  expect(await controller.removeComment(99)).toEqual({ ok: false, failure: { code: 'command-failed', message: 'unknown-comment' } });
  expect(() => controller.ref.resolveComment(99)).not.toThrow();
  expect(await controller.ref.insertCommentReply(7.5, 'Reply', 'B')).toBeNull();
  expect(submit).not.toHaveBeenCalled();
  expect(changes).toEqual([]);
});

for (const [options, code] of [[{ readOnly: true }, 'read-only'], [{ mode: 'viewing' }, 'viewing-mode']] as const) {
  test(`a ${code} editor refuses comment changes as values and still reads comments`, async () => {
    const { controller, submit, host } = reviewFixture(options);
    expect(await controller.ref.insertCommentReply(7, 'Reply', 'B')).toBeNull();
    expect(await controller.setCommentDone(7, true)).toMatchObject({ ok: false, failure: { code } });
    expect(await controller.removeComment(7)).toMatchObject({ ok: false, failure: { code } });
    expect(() => controller.ref.resolveComment(7)).not.toThrow();
    expect(submit).not.toHaveBeenCalled();
    expect(controller.ref.getComments()).toEqual(host);
  });
}

test('getComments follows the latest review read of the worker document', async () => {
  const { controller, host } = reviewFixture();
  host.length = 0;
  expect(controller.ref.getComments()).toEqual([]);
  await controller.query({ kind: 'reviewSnapshot' });
  expect(controller.ref.getComments()).toEqual([{ id: 7, author: 'A', content: [] }]);
});

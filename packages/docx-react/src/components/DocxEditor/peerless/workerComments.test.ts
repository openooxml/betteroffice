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
    if (intent.kind === 'removeComment') comments = comments.filter(comment => comment.id !== 7);
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
  return { controller, binding, viewStore, read, submit, ranges, host, changes, nextChange: () => new Promise<void>(resolve => { changed = resolve; }) };
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


test('removing a worker comment admits its native id and emits the shortened projection once', async () => {
  const { controller, submit, changes } = reviewFixture();
  expect(await controller.removeComment(7)).toEqual({ ok: true, status: 'executed' });
  expect(submit).toHaveBeenCalledTimes(1);
  expect(submit.mock.calls[0]![0]).toEqual({ kind: 'removeComment', payload: { id: 'native-7', expectVersion: 'native' } });
  expect(controller.ref.getComments()).toEqual([]);
  expect(changes).toEqual([[]]);
});

for (const barrier of ['save', 'flushPendingInput'] as const) {
  test(`${barrier} joins comment operations queued before admission`, async () => {
    const { controller, read, submit } = reviewFixture();
    const original = read.getMockImplementation()!;
    let release!: () => void;
    const held = new Promise<void>(resolve => { release = resolve; });
    let reading!: () => void;
    const started = new Promise<void>(resolve => { reading = resolve; });
    read.mockImplementation(async request => {
      if (request.kind === 'reviewTarget') { reading(); await held; }
      return original(request);
    });
    controller.ref.resolveComment(7);
    let completed = false;
    const pending = controller.ref[barrier]().then(() => { completed = true; });
    await started;
    expect(completed).toBe(false);
    release();
    await pending;
    expect(submit).toHaveBeenCalledTimes(1);
    expect(controller.ref.getComments()[0]?.done).toBe(true);
  });
}

test('getComments preserves its projection through ordinary versions and accepts a lower version in a new incarnation', async () => {
  const { controller, viewStore, host } = reviewFixture();
  await controller.setCommentDone(7, true);
  const previous = viewStore.getSnapshot();
  viewStore.acceptWorkerState(previous.identity, { epoch: 1, version: 10, throughSeq: 2, presentedThroughSeq: 0, selection: null, editState });
  expect(controller.ref.getComments()).toEqual([{ ...host[0]!, done: true }]);
  const identity = { ...previous.identity, incarnation: 2, epoch: 2 };
  viewStore.acceptWorkerState(identity, { epoch: 2, version: 0, throughSeq: 0, presentedThroughSeq: 0, selection: null, editState });
  viewStore.installFrameGeometry({ identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] });
  expect(controller.ref.getComments()).toEqual(host);
  await controller.query({ kind: 'reviewSnapshot' });
  expect(controller.ref.getComments()).toEqual([{ ...host[0]!, done: true }]);
});


test('a review refresh from the previous incarnation cannot overwrite getComments', async () => {
  const { controller, read, viewStore, host } = reviewFixture();
  const original = read.getMockImplementation()!;
  let release!: () => void;
  const held = new Promise<void>(resolve => { release = resolve; });
  let reading!: () => void;
  const started = new Promise<void>(resolve => { reading = resolve; });
  read.mockImplementation(async request => {
    const result = await original(request);
    if (request.kind === 'reviewSnapshot') { reading(); await held; }
    return result;
  });
  const pending = controller.query({ kind: 'reviewSnapshot' });
  await started;
  const identity = { session: 's', incarnation: 2, epoch: 2 };
  viewStore.acceptWorkerState(identity, { epoch: 2, version: 0, throughSeq: 0, presentedThroughSeq: 0, selection: null, editState });
  viewStore.installFrameGeometry({ identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] });
  release();
  await expect(pending).rejects.toThrow('Worker read was superseded');
  expect(controller.ref.getComments()).toEqual(host);
  await controller.query({ kind: 'reviewSnapshot' });
  expect(controller.ref.getComments()).toEqual([{ id: 7, author: 'A', content: [] }]);
});


test('a new incarnation starts its own presented review refresh before an old read settles', async () => {
  const { controller, read, viewStore } = reviewFixture();
  const original = read.getMockImplementation()!;
  let release!: () => void;
  const held = new Promise<void>(resolve => { release = resolve; });
  let reading!: () => void;
  const started = new Promise<void>(resolve => { reading = resolve; });
  let first = true;
  read.mockImplementation(async request => {
    const response = await original(request);
    if (request.kind !== 'reviewSnapshot') return response;
    const old = first;
    first = false;
    if (old) { reading(); await held; }
    return { ...response, value: { version: 'native', previewVersion: 0, frameEpoch: 1, layoutRevision: 1,
      comments: [], uiComments: [{ id: old ? 70 : 80, author: 'B', content: [] }], commentIds: {},
      proposals: { version: 'native', previewVersion: 0, proposals: [] },
      proposalGeometry: { version: 'native', previewVersion: 0 } } as never };
  });
  const install = (identity: { session: string; incarnation: number; epoch: number }) => viewStore.installFrameGeometry({
    identity, documentVersion: 'native', coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [],
  });
  install(viewStore.getSnapshot().identity);
  const old = controller.review();
  void old.catch(() => {});
  await started;
  const fresh = controller.review();
  const identity = { session: 's', incarnation: 2, epoch: 2 };
  viewStore.acceptWorkerState(identity, { epoch: 2, version: 0, throughSeq: 0, presentedThroughSeq: 0, selection: null, editState });
  install(identity);
  try {
    expect((await fresh).uiComments).toEqual([{ id: 80, author: 'B', content: [] }]);
    await expect(old).rejects.toThrow('Document was replaced');
    release();
    expect(controller.ref.getComments()).toEqual([{ id: 80, author: 'B', content: [] }]);
    expect((await controller.review()).uiComments).toEqual([{ id: 80, author: 'B', content: [] }]);
  } finally { release(); controller.dispose(); viewStore.dispose(); }
});

type Note = { id: number; author: string; date: string; content: unknown[]; done?: boolean; parentId?: number };
const note = (id: number, text: string, extra: Partial<Note> = {}): Note => ({ id, author: 'A', date: '2026-01-01T00:00:00Z',
  content: [{ type: 'paragraph', formatting: {}, content: [{ type: 'run', formatting: {}, content: [{ type: 'text', text }] }] }], ...extra });

function sidebarFixture(options: { hostComments?: Note[]; readOnly?: boolean; mode?: 'viewing' } = {}) {
  const range = { story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 7 } };
  const selection = { anchor: { story: 'body', paraId: 'p', offset: 7 }, focus: { story: 'body', paraId: 'p', offset: 3 } };
  let anchored = range;
  let projected: Note[] = [note(7, 'Root'), note(8, 'Child', { parentId: 7 })];
  let seq = 0;
  let version = 0;
  const viewStore = createPeerlessViewStore({ identity: { session: 's', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: selection, anchors: null, editState, geometry: null });
  const read = mock(async (request: { kind: string; target?: { id: number }; selection?: unknown }) => ({ status: 'ok', version: `native-${version}`, value:
    request.kind === 'readParagraphs' ? { ok: true, version: `native-${version}`, paragraphs: [] } :
    request.kind === 'commentSelectionRange' ? anchored :
    request.kind === 'reviewTarget' ? projected.some(comment => comment.id === request.target!.id) ? { id: `n${request.target!.id}`, ranges: [range] } : null :
    request.kind === 'reviewSnapshot' ? { comments: [], uiComments: structuredClone(projected), commentIds: {}, frameEpoch: 1, layoutRevision: 1 } : null }));
  const submit = mock((intent: { kind: string; payload: { id?: string; body?: string; parentId?: string; patch?: { done: boolean }; ranges?: unknown; author?: string; expectVersion?: string } }) => {
    seq += 1;
    const stale = intent.payload.expectVersion !== undefined && intent.payload.expectVersion !== `native-${version}`;
    const id = Number(intent.payload.id?.slice(1));
    if (!stale) {
      version += 1;
      if (intent.kind === 'addComment') projected = [...projected, note(20 + seq, intent.payload.body!, intent.payload.parentId ? { parentId: Number(intent.payload.parentId.slice(1)) } : {})];
      if (intent.kind === 'updateComment') projected = projected.map(comment => comment.id === id ? { ...comment, done: intent.payload.patch!.done } : comment);
      if (intent.kind === 'removeComment') projected = projected.filter(comment => comment.id !== id);
    }
    const receipt = stale ? {} : intent.kind === 'addComment' ? { comment: { seq, id: 20 + seq, comments: structuredClone(projected) } } : { commentChange: { seq, comments: structuredClone(projected) } };
    const view = viewStore.getSnapshot();
    viewStore.acceptWorkerState(view.identity, { epoch: 1, version: view.version + (stale ? 0 : 1), throughSeq: seq, presentedThroughSeq: 0, selection, editState: { ...editState, ...receipt } as never });
    return { status: 'submitted', ticket: { seq, ack: Promise.resolve(stale ? { status: 'refused', seq, reason: 'stale-version' } : { status: 'applied', seq }), presented: Promise.resolve('presented') } };
  });
  const events: unknown[][] = [];
  const settings = { read: read as never, readAt: read as never, focus() {}, zoom: 1, author: 'Me',
    readOnly: options.readOnly, mode: options.mode, hostComments: options.hostComments as never,
    onCommentsChange: (comments: unknown) => events.push(['change', comments]), onCommentAdd: (comment: unknown) => events.push(['add', comment]),
    onCommentReply: (reply: unknown, parent: unknown) => events.push(['reply', reply, parent]), onCommentResolve: (comment: unknown) => events.push(['resolve', comment]),
    onCommentDelete: (comment: unknown) => events.push(['delete', comment]) };
  const controller = createPeerlessShellController(settings);
  const binding = { session: { flush: async () => {} }, admission: {}, commands: { submit }, viewStore } as unknown as PeerlessEditSessionBinding;
  controller.attach(binding);
  return { controller, submit, read, events, range, selection, settings, projected: () => projected,
    anchor: (next: typeof range) => { anchored = next; } };
}

test('comment sidebar actions apply through the worker and notify the host in hybrid order', async () => {
  const { controller, submit, events, range, selection, projected } = sidebarFixture();
  await controller.query({ kind: 'reviewSnapshot' });
  const draft = await controller.beginComment();
  expect(draft).toMatchObject({ selection });
  expect(events).toEqual([]);
  expect(await controller.comment({ kind: 'add', draft: draft!, text: 'New' })).toEqual({ ok: true, status: 'executed' });
  expect(submit.mock.calls[0]![0]).toMatchObject({ kind: 'addComment', payload: { ranges: [range], author: 'Me', body: 'New', expectVersion: 'native-0' } });
  const added = projected().at(-1)!;
  expect(events.splice(0)).toEqual([['change', projected()], ['add', added]]);
  const root = projected()[0]!;
  expect(await controller.comment({ kind: 'reply', id: 7, text: 'Answer' })).toMatchObject({ ok: true });
  expect(submit.mock.calls[1]![0]).toMatchObject({ kind: 'addComment', payload: { parentId: 'n7', body: 'Answer', author: 'Me' } });
  expect(events.splice(0)).toEqual([['change', projected()], ['reply', projected().at(-1), root]]);
  expect(await controller.comment({ kind: 'resolve', id: 7 })).toMatchObject({ ok: true });
  expect(events.splice(0)).toEqual([['change', projected()], ['resolve', { ...root, done: true }]]);
  expect(await controller.comment({ kind: 'reopen', id: 7 })).toMatchObject({ ok: true });
  expect(events.splice(0)).toEqual([['change', projected()]]);
  const reopened = projected()[0]!;
  expect(await controller.comment({ kind: 'delete', id: 7 })).toMatchObject({ ok: true });
  expect(submit.mock.calls.slice(4).map(call => call[0])).toEqual([
    { kind: 'removeComment', payload: { id: 'n8', expectVersion: 'native-4' } },
    { kind: 'removeComment', payload: { id: 'n22', expectVersion: 'native-5' } },
    { kind: 'removeComment', payload: { id: 'n7', expectVersion: 'native-6' } }]);
  expect(projected().map(comment => comment.id)).toEqual([21]);
  expect(events).toEqual([['change', projected()], ['delete', reopened]]);
});

test('read-only and viewing editors refuse comment sidebar actions without notifying the host', async () => {
  for (const options of [{ readOnly: true }, { mode: 'viewing' as const }]) {
    const { controller, submit, events } = sidebarFixture(options);
    await controller.query({ kind: 'reviewSnapshot' });
    expect(await controller.beginComment()).toBeNull();
    for (const action of [{ kind: 'reply', id: 7, text: 'No' }, { kind: 'resolve', id: 7 }, { kind: 'reopen', id: 7 }, { kind: 'delete', id: 7 }] as const) {
      expect(await controller.comment(action)).toMatchObject({ ok: false });
    }
    expect(submit).not.toHaveBeenCalled();
    expect(events).toEqual([]);
  }
});

test('controlled comment actions change only the requested entries of the host list', async () => {
  const hosted = [note(7, 'Host root'), note(30, 'Host only')];
  const { controller, submit, events } = sidebarFixture({ hostComments: hosted });
  expect(await controller.comment({ kind: 'resolve', id: 7 })).toMatchObject({ ok: true });
  expect(submit).toHaveBeenCalledTimes(1);
  expect(events.splice(0)).toEqual([['change', [{ ...hosted[0], done: true }, hosted[1]]], ['resolve', { ...hosted[0], done: true }]]);
  expect(await controller.comment({ kind: 'resolve', id: 30 })).toMatchObject({ ok: true });
  expect(await controller.comment({ kind: 'reopen', id: 30 })).toMatchObject({ ok: true });
  expect(submit).toHaveBeenCalledTimes(1);
  expect(events.splice(0)).toEqual([['change', [hosted[0], { ...hosted[1], done: true }]], ['resolve', { ...hosted[1], done: true }],
    ['change', [hosted[0], { ...hosted[1], done: undefined }]]]);
  expect(await controller.comment({ kind: 'reply', id: 30, text: 'Metadata reply' })).toMatchObject({ ok: true });
  const [[, changed], [, reply, parent]] = events.splice(0) as [[string, Note[]], [string, Note, Note]];
  expect(reply).toMatchObject({ id: 31, parentId: 30, author: 'Me' });
  expect(changed).toEqual([...hosted, reply]);
  expect(parent).toEqual(hosted[1]);
  expect(await controller.comment({ kind: 'delete', id: 30 })).toMatchObject({ ok: true });
  expect(submit).toHaveBeenCalledTimes(1);
  expect(events).toEqual([['change', [hosted[0]]], ['delete', hosted[1]]]);
  expect(controller.ref.getComments()).toEqual(hosted as never);
});

test('a composer draft re-reads its anchor when another change lands first and adds at the current version', async () => {
  const { controller, submit, read, events, range, selection, anchor } = sidebarFixture();
  await controller.query({ kind: 'reviewSnapshot' });
  const draft = await controller.beginComment();
  expect(await controller.comment({ kind: 'resolve', id: 7 })).toMatchObject({ ok: true });
  events.splice(0);
  const moved = { ...range, start: { paraId: 'p', offset: 5 }, end: { paraId: 'p', offset: 9 } };
  anchor(moved);
  expect(await controller.comment({ kind: 'add', draft: draft!, text: 'Late' })).toEqual({ ok: true, status: 'executed' });
  expect(read.mock.calls.filter(([request]) => request.kind === 'commentSelectionRange').at(-1)![0]).toEqual({ kind: 'commentSelectionRange', selection });
  expect(submit.mock.calls.at(-1)![0]).toMatchObject({ kind: 'addComment', payload: { ranges: [moved], body: 'Late', expectVersion: 'native-1' } });
  expect(events.map(([kind]) => kind)).toEqual(['change', 'add']);
});

test('a composer add whose anchored text is gone refuses as a value without admission', async () => {
  const { controller, submit, events, anchor } = sidebarFixture();
  const draft = await controller.beginComment();
  anchor({ story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 3 } });
  expect(await controller.comment({ kind: 'add', draft: draft!, text: 'Gone' }))
    .toEqual({ ok: false, failure: { code: 'command-failed', message: 'invalid-range' } });
  expect(submit).not.toHaveBeenCalled();
  expect(events).toEqual([]);
});

test('a composer open that turns read-only while it reads the selection opens nothing', async () => {
  const { controller, settings } = sidebarFixture();
  const pending = controller.beginComment();
  controller.update({ ...settings, readOnly: true });
  expect(await pending).toBeNull();
});

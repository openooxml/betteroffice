import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useCommentSidebarItems } from '../../../hooks/useCommentSidebarItems';
import { useWorkerComments } from './useWorkerComments';
import { createPeerlessShellController } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import type { DocxCommandResult } from '../../../commands/types';
import type { WorkerCommentAction, WorkerCommentDraft } from './shellController';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const refused = (message: string): DocxCommandResult => ({ ok: false, failure: { code: 'command-failed', message } });

function mount(worker?: ReturnType<typeof createPeerlessShellController>) {
  const draft = { binding: {}, selection: { anchor: { story: 'body', paraId: 'p', offset: 1 }, focus: { story: 'body', paraId: 'p', offset: 4 } } } as unknown as WorkerCommentDraft;
  const results: DocxCommandResult[] = [];
  const comment = mock(async (_action: WorkerCommentAction): Promise<DocxCommandResult> => results.shift() ?? { ok: true, status: 'executed' });
  const controller = worker ?? { beginComment: mock(async () => draft), comment } as never;
  const errors: string[] = [];
  let hook!: ReturnType<typeof useWorkerComments>;
  function Harness({ readOnly }: { readOnly: boolean }) {
    hook = useWorkerComments({ controller, documentKey: 'doc', readOnly, setOpen() {}, setExpanded() {},
      selectionAnchor: () => ({ y: 10, right: 0 }), focus() {}, onError: error => errors.push(error.message) });
    const items = useCommentSidebarItems({ comments: [], trackedChanges: [], callbacks: hook.callbacks,
      isAddingComment: hook.adding, addCommentYPosition: hook.addCommentY });
    return <>{items.map(item => <div key={item.id}>{item.render({ isExpanded: false, onToggleExpand() {}, measureRef() {} })}</div>)}</>;
  }
  const view = render(<Harness readOnly={false} />);
  const textarea = () => view.container.querySelector('textarea');
  const submit = () => [...view.container.querySelectorAll('button')].at(-1)!;
  return { view, Harness, hook: () => hook, comment, results, errors, draft, textarea, submit };
}

function heldWorker() {
  const selection = { anchor: { story: 'body', paraId: 'p', offset: 1 }, focus: { story: 'body', paraId: 'p', offset: 4 } };
  const editState = { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null, canUndo: false, canRedo: false, hyperlink: null, tableContext: null };
  const identity = { session: 's', incarnation: 1, epoch: 1 };
  const viewStore = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0, workerSelection: selection, anchors: null, editState, geometry: null });
  const read = async (request: { kind: string }) => ({ status: 'ok', version: 'native', value:
    request.kind === 'readParagraphs' ? { ok: true, version: 'native', paragraphs: [] } :
    request.kind === 'commentSelectionRange' ? { story: 'body', start: { paraId: 'p', offset: 1 }, end: { paraId: 'p', offset: 4 } } :
    request.kind === 'reviewSnapshot' ? { comments: [], uiComments: [], commentIds: {}, frameEpoch: 1, layoutRevision: 1 } : null });
  const acks: ((outcome: { status: 'applied' | 'refused'; seq: number; reason?: string }) => void)[] = [];
  const submit = mock((intent: { payload: { body: string } }) => {
    const seq = acks.length + 1;
    const comments = [{ id: 20 + seq, author: 'Me', date: '2026-01-01T00:00:00Z', content: [{ type: 'paragraph', formatting: {}, content: [{ type: 'run', formatting: {}, content: [{ type: 'text', text: intent.payload.body }] }] }] }];
    viewStore.acceptWorkerState(identity, { epoch: 1, version: seq, throughSeq: seq, presentedThroughSeq: 0, selection, editState: { ...editState, comment: { seq, id: 20 + seq, comments } } as never });
    return { status: 'submitted', ticket: { seq, ack: new Promise(resolve => { acks.push(resolve); }), presented: Promise.resolve('presented') } };
  });
  const controller = createPeerlessShellController({ read: read as never, readAt: read as never, focus() {}, zoom: 1, author: 'Me' });
  controller.attach({ session: { flush: async () => {} }, admission: {}, commands: { submit }, viewStore } as unknown as PeerlessEditSessionBinding);
  return { controller, submit, acks };
}

const settle = () => act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });

test('the composer is read-only while its add waits for the worker acknowledgement', async () => {
  const worker = heldWorker();
  const { hook, errors, textarea, submit } = mount(worker.controller);
  await act(async () => { await hook().beginAddComment(); });
  fireEvent.change(textarea()!, { target: { value: 'A' } });
  await act(async () => { fireEvent.click(submit()); });
  await settle();
  expect(worker.submit.mock.calls.map(([intent]) => intent.payload.body)).toEqual(['A']);
  expect(textarea()!.readOnly).toBe(true);
  await act(async () => { worker.acks[0]!({ status: 'refused', seq: 1, reason: 'stale-version' }); });
  await settle();
  expect(errors).toEqual(['stale-version']);
  expect(textarea()!.readOnly).toBe(false);
  expect(textarea()!.value).toBe('A');
  fireEvent.change(textarea()!, { target: { value: 'AB' } });
  await act(async () => { fireEvent.click(submit()); });
  await settle();
  expect(worker.submit.mock.calls.map(([intent]) => intent.payload.body)).toEqual(['A', 'AB']);
  expect(textarea()!.readOnly).toBe(true);
  await act(async () => { worker.acks[1]!({ status: 'applied', seq: 2 }); });
  await settle();
  expect(textarea()).toBeNull();
  expect(hook().adding).toBe(false);
});

test('a refused composer add keeps the composer and its text and reports the refusal; an executed add closes it', async () => {
  const { hook, comment, results, errors, draft, textarea, submit } = mount();
  await act(async () => { await hook().beginAddComment(); });
  fireEvent.change(textarea()!, { target: { value: 'Note' } });
  results.push(refused('stale-version'));
  await act(async () => { fireEvent.click(submit()); });
  expect(comment.mock.calls.at(-1)![0]).toEqual({ kind: 'add', draft, text: 'Note' });
  expect(errors).toEqual(['stale-version']);
  expect(textarea()?.value).toBe('Note');
  expect(hook().adding).toBe(true);
  await act(async () => { fireEvent.click(submit()); });
  expect(comment).toHaveBeenCalledTimes(2);
  expect(textarea()).toBeNull();
  expect(hook().adding).toBe(false);
  expect(errors).toEqual(['stale-version']);
});

test('refused sidebar replies, resolves and deletes report through onError', async () => {
  const { hook, results, errors } = mount();
  results.push(refused('unknown-comment'), refused('read-only'), refused('document-replaced'));
  await act(async () => {
    hook().callbacks.onCommentReply!(7, 'Answer');
    hook().callbacks.onCommentResolve!(7);
    hook().callbacks.onCommentDelete!(7);
    await new Promise<void>((resolve) => setImmediate(resolve));
  });
  expect(errors).toEqual(['unknown-comment', 'read-only', 'document-replaced']);
});

test('an open composer hides while the editor is read-only and keeps its draft and text for editing again', async () => {
  const { view, Harness, hook, comment, textarea } = mount();
  await act(async () => { await hook().beginAddComment(); });
  fireEvent.change(textarea()!, { target: { value: 'Kept' } });
  view.rerender(<Harness readOnly />);
  expect(hook().adding).toBe(true);
  expect(hook().composerHidden).toBe(true);
  expect(textarea()?.value).toBe('Kept');
  view.rerender(<Harness readOnly={false} />);
  expect(hook().composerHidden).toBe(false);
  expect(textarea()?.value).toBe('Kept');
  expect(comment).not.toHaveBeenCalled();
});

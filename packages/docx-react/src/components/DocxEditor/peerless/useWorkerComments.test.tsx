import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useCommentSidebarItems } from '../../../hooks/useCommentSidebarItems';
import { useWorkerComments } from './useWorkerComments';
import type { DocxCommandResult } from '../../../commands/types';
import type { WorkerCommentAction, WorkerCommentDraft } from './shellController';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const refused = (message: string): DocxCommandResult => ({ ok: false, failure: { code: 'command-failed', message } });

function mount() {
  const draft = { binding: {}, selection: { anchor: { story: 'body', paraId: 'p', offset: 1 }, focus: { story: 'body', paraId: 'p', offset: 4 } } } as unknown as WorkerCommentDraft;
  const results: DocxCommandResult[] = [];
  const comment = mock(async (_action: WorkerCommentAction): Promise<DocxCommandResult> => results.shift() ?? { ok: true, status: 'executed' });
  const controller = { beginComment: mock(async () => draft), comment } as never;
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

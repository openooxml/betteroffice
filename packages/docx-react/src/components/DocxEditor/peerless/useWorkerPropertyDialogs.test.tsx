import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useWorkerPropertyDialogs } from './useWorkerPropertyDialogs';
import type { WorkerShellController } from './workerProperties';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

function fixture() {
  const selection = { anchor: { story: 'body', paraId: 'p', offset: 1 }, focus: { story: 'body', paraId: 'p', offset: 4 } };
  const state = { status: 'ready', view: { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1, workerSelection: selection, geometry: null } };
  const applyQuery = mock(async (_query: unknown, _intent: unknown) => ({ ok: true, status: 'executed' }));
  const readAt = mock(async () => ({ selectedText: 'Old' }));
  const controller = { subscribe: () => () => {}, getSnapshot: () => state, applyQuery, readAt,
    environment: () => ({ readOnly: false, mode: 'editing' }), ref: { focus() {}, flushPendingInput: async () => {} },
    query: async (request: { target?: { kind: string } }) => ({ version: 'captured-native', selection,
      value: request.target?.kind === 'object' ? { kind: 'object', handle: 'physical-second', properties: { width: 80, height: 40, alt: 'Old' } } :
        { kind: 'hyperlink', hyperlink: { href: 'https://example.com/old', tooltip: null } } }),
  } as unknown as WorkerShellController;
  const hook = renderHook(() => useWorkerPropertyDialogs(controller, {}));
  return { hook, applyQuery, selection, readAt };
}

test('hyperlink edits use the opened worker selection and expose its display text as read-only', async () => {
  const { hook, applyQuery, selection, readAt } = fixture();
  await act(async () => { await hook.result.current.openLink(); });
  const dialog = hook.result.current.element.props.children[2];
  expect(dialog.props.displayTextReadOnly).toBe(true);
  expect(readAt).toHaveBeenCalledWith({ kind: 'rawSelectionInfo', selection }, 'captured-native');
  await act(async () => { await dialog.props.onSubmit({ url: 'https://example.com/new', tooltip: 'New' }); });
  expect(applyQuery.mock.calls[0]![1]).toEqual({ kind: 'setHyperlink', payload: { selection,
    hyperlink: { href: 'https://example.com/new', tooltip: 'New' }, expectVersion: 'captured-native' } });
});

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useWorkerHeaderFooterEditing } from './useHeaderFooterEditing';
import type { createPeerlessShellController } from '../peerless/shellController';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

function fixture() {
  const state = { status: 'ready', view: { editState: { activePart: { kind: 'footer', story: 'hf:shared', section: 'second', page: 3, variant: 'even' } } } };
  const query = mock(async (request: { kind: string }) => ({ version: 'native', value: request.kind === 'paragraphIdentities' ?
    { paragraphs: [{ session: { story: 'body', paraId: 'body-start' } }] } : { metadata: { body: {} } } }));
  const applyQuery = mock(async () => ({ ok: true, status: 'executed' }));
  const focus = mock(() => {});
  const onError = mock(() => {});
  const submit = mock(async () => ({ ok: true, status: 'noop' }));
  const controller = { subscribe: () => () => {}, getSnapshot: () => state, query, applyQuery, submit, ref: { focus } } as unknown as ReturnType<typeof createPeerlessShellController>;
  const hook = renderHook(() => useWorkerHeaderFooterEditing(controller, onError));
  return { hook, query, applyQuery, focus, onError };
}

test('worker parts preserve section page variant and note targets and omit story from removal', async () => {
  const { hook, applyQuery, focus } = fixture();
  await act(async () => {
    await hook.result.current.ensure({ kind: 'header', section: 'second', page: 3, variant: 'even' });
    await hook.result.current.open({ kind: 'footnote', section: 'second', page: 3, noteId: 8 });
    await hook.result.current.remove();
  });
  expect(applyQuery.mock.calls.map(call => (call as unknown[])[1])).toEqual([
    { kind: 'ensurePart', payload: { kind: 'header', section: 'second', page: 3, variant: 'even', expectVersion: 'native' } },
    { kind: 'openPart', payload: { kind: 'footnote', section: 'second', page: 3, noteId: 8, expectVersion: 'native' } },
    { kind: 'removePart', payload: { kind: 'footer', section: 'second', page: 3, variant: 'even', noteId: undefined, expectVersion: 'native' } },
  ]);
  expect(focus).toHaveBeenCalledTimes(3);
});

test('closing a part selects a worker body identity and reports refused writes', async () => {
  const { hook, applyQuery, onError } = fixture();
  await act(async () => { await hook.result.current.close(); });
  expect((applyQuery.mock.calls[0] as unknown[])[1]).toEqual({ kind: 'select', payload: { selection: {
    anchor: { story: 'body', paraId: 'body-start', offset: 0 }, focus: { story: 'body', paraId: 'body-start', offset: 0 },
  } } });
  applyQuery.mockImplementation(async () => ({ ok: false, failure: { message: 'stale-version' } }) as never);
  await act(async () => { await hook.result.current.remove(); });
  expect(onError).toHaveBeenCalledWith(expect.objectContaining({ message: 'stale-version' }));
});

test('note entry awaits the point command and uses its native story and selection', async () => {
  const { hook, query, applyQuery, onError } = fixture();
  const at = { story: 'fn:2', paraId: 'second-note', offset: 4 };
  query.mockImplementation(async () => ({ version: 'native', value: { metadata: { body: {} } }, selection: { anchor: at, focus: at } }) as never);
  const point = { page: { pageId: '1', pageIndex: 0 }, x: 100, y: 755, frameEpoch: 1, layoutRevision: 1,
    clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false } };
  await act(async () => { await hook.result.current.openAtPoint(point, 'final'); });
  expect(applyQuery.mock.calls.map(call => (call as unknown[])[1])).toEqual([
    { kind: 'openPart', payload: { kind: 'footnote', noteId: 2, section: 'final', page: 0, expectVersion: 'native' } },
    { kind: 'select', payload: { selection: { anchor: at, focus: at } } },
  ]);
  expect(onError).not.toHaveBeenCalled();
});

test('part entry resolves layout section ordinals against source metadata', async () => {
  const { hook, query, applyQuery } = fixture();
  query.mockImplementation(async () => ({ version: 'native', value: { metadata: { body: { sections: [{ id: 'authored-first' }, { id: 'authored-last' }] } } } }) as never);
  await act(async () => {
    await hook.result.current.ensure({ kind: 'header', section: 'layout-first', page: 0 }, 0);
    await hook.result.current.ensure({ kind: 'footer', section: 'layout-last', page: 2 }, 1);
  });
  expect(applyQuery.mock.calls.map(call => (call as unknown[])[1])).toEqual([
    { kind: 'ensurePart', payload: { kind: 'header', section: 'authored-first', page: 0, expectVersion: 'native' } },
    { kind: 'ensurePart', payload: { kind: 'footer', section: 'final', page: 2, expectVersion: 'native' } },
  ]);
});

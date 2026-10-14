import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useWorkerContextMenu } from './useWorkerContextMenu';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, fireEvent } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

test('image context resize targets the clicked physical handle rather than the previous caret', async () => {
  const surface = document.createElement('div');
  surface.className = 'paged-editor__page';
  surface.dataset.pageIndex = '0';
  surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 200, height: 200 } as DOMRect);
  document.body.appendChild(surface);
  const snapshot = { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1,
    workerSelection: null, geometry: { frameEpoch: 2, layoutRevision: 3, documentVersion: 'native',
      pages: [{ page: { pageId: 'p', pageIndex: 0 }, bounds: { x: 0, y: 0, width: 200, height: 200 } }],
      objects: { version: 'native', frameEpoch: 2, layoutRevision: 3, objects: [
        { kind: 'image', handle: 'first', story: 'body', rects: [{ pageIndex: 0, x: 0, y: 0, width: 40, height: 20 }] },
        { kind: 'image', handle: 'second', story: 'body', rects: [{ pageIndex: 0, x: 60, y: 0, width: 40, height: 20 }] },
      ] } } };
  const submit = mock(async (_intent: unknown) => ({ ok: true }));
  const binding = { viewStore: { getSnapshot: () => snapshot } };
  let menu!: ReturnType<typeof useWorkerContextMenu>;
  function Host() {
    menu = useWorkerContextMenu({ binding: () => binding,
      environment: () => ({ readOnly: false, mode: 'editing', translate: (key: string) => key }),
      commands: { getState: () => ({ enabled: false }) }, clipboard: async () => {}, submit,
    } as unknown as Parameters<typeof useWorkerContextMenu>[0]);
    return menu.element;
  }
  const view = render(<Host />);
  await act(async () => { menu.open({ target: surface, clientX: 80, clientY: 10, preventDefault() {}, stopPropagation() {} } as never); });
  fireEvent.click(view.getByRole('menuitem', { name: 'Resize image' }));
  const dialog = view.getByRole('dialog', { name: 'Resize image' });
  fireEvent.change(view.getByLabelText('Width (px)'), { target: { value: '80' } });
  fireEvent.change(view.getByLabelText('Height (px)'), { target: { value: '40' } });
  await act(async () => { fireEvent.click(dialog.querySelector('button[type=submit]')!); });
  expect(submit.mock.calls[0]![0]).toEqual({ kind: 'resizeObject', payload: { handle: 'second', expectVersion: 'native', widthEmu: 762000, heightEmu: 381000 } });
  surface.remove();
});

test.each(['edited', 'replaced', 'readOnly'] as const)('object dialog refuses a target after it becomes %s', async reason => {
  const surface = document.createElement('div');
  surface.className = 'paged-editor__page';
  surface.dataset.pageIndex = '0';
  surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 400, height: 400 } as DOMRect);
  document.body.appendChild(surface);
  const snapshot = { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1,
    workerSelection: null, geometry: { frameEpoch: 2, layoutRevision: 3, documentVersion: 'native',
      pages: [{ page: { pageId: 'p', pageIndex: 0 }, bounds: { width: 200, height: 200 } }],
      objects: { version: 'native', frameEpoch: 2, layoutRevision: 3, objects: [
        { kind: 'image', handle: 'image', rects: [{ pageIndex: 0, x: 60, y: 0, width: 40, height: 20 }] },
      ] } } };
  let binding = { viewStore: { getSnapshot: () => snapshot } };
  let readOnly = false;
  const submit = mock(async (_intent: unknown) => ({ ok: true }));
  const onError = mock((_error: Error) => {});
  let menu!: ReturnType<typeof useWorkerContextMenu>;
  function Host() {
    menu = useWorkerContextMenu({ binding: () => binding,
      environment: () => ({ readOnly, mode: 'editing', translate: (key: string) => key }),
      commands: { getState: () => ({ enabled: false }) }, clipboard: async () => {}, submit, onError,
    } as unknown as Parameters<typeof useWorkerContextMenu>[0]);
    return menu.element;
  }
  const view = render(<Host />);
  await act(async () => { menu.open({ target: surface, clientX: 160, clientY: 20, preventDefault() {}, stopPropagation() {} } as never); });
  fireEvent.click(view.getByRole('menuitem', { name: 'Resize image' }));
  if (reason === 'edited') snapshot.version += 1;
  else if (reason === 'replaced') binding = { viewStore: binding.viewStore };
  else readOnly = true;
  const dialog = view.getByRole('dialog', { name: 'Resize image' });
  await act(async () => { fireEvent.click(dialog.querySelector('button[type=submit]')!); });
  expect(submit).not.toHaveBeenCalled();
  expect(onError).toHaveBeenCalledTimes(1);
  expect(view.getByRole('dialog', { name: 'Resize image' })).toBeDefined();
  surface.remove();
});

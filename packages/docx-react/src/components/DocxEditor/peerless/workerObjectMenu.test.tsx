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

const range = (story: string, from: number, to: number, sticky: number) => ({
  anchor: { story, paraId: `${story}:p0`, offset: from, sticky: new Uint8Array([0, sticky, from]) },
  focus: { story, paraId: `${story}:p0`, offset: to, sticky: new Uint8Array([0, sticky, to]) },
});
const inTable = (tableIndex: number, rowIndex: number) => ({ tableContext: { story: 'body', tableIndex, rowIndex, columnIndex: 0,
  rowCount: 2, columnCount: 2, hasMultiCellSelection: false, canSplitCell: false } });

async function clickAfterRepublish(label: string, republish: (snapshot: Record<string, unknown>) => void) {
  const snapshot: Record<string, unknown> = { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1,
    workerSelection: range('body:t0:r0c0', 0, 2, 1), editState: inTable(0, 0), geometry: null };
  const submit = mock((_command: unknown) => {});
  const execute = mock(async (..._args: unknown[]) => ({ ok: true }));
  const clipboard = mock(async (_cut: boolean) => {});
  const onError = mock((_error: Error) => {});
  const binding = { viewStore: { getSnapshot: () => snapshot }, commands: { submit } };
  let menu!: ReturnType<typeof useWorkerContextMenu>;
  function Host() {
    menu = useWorkerContextMenu({ binding: () => binding,
      environment: () => ({ readOnly: false, mode: 'editing', translate: (key: string) => key }),
      commands: { getState: () => ({ enabled: true }), execute }, clipboard, onError,
    } as unknown as Parameters<typeof useWorkerContextMenu>[0]);
    return menu.element;
  }
  const view = render(<Host />);
  await act(async () => { menu.open({ target: document.body, clientX: 0, clientY: 0, preventDefault() {}, stopPropagation() {} } as never); });
  republish(snapshot);
  await act(async () => { fireEvent.click(view.getByRole('menuitem', { name: label })); });
  return { submit, execute, clipboard, onError };
}

test('context menu acts after a version bump and a sticky-only selection republish', async () => {
  const table = await clickAfterRepublish('table.selectTable', snapshot => Object.assign(snapshot,
    { version: 2, workerSelection: range('body:t0:r1c0', 0, 2, 9), editState: inTable(0, 1) }));
  expect(table.onError).not.toHaveBeenCalled();
  expect(table.execute).toHaveBeenCalledWith('tableAction', 'selectTable');
  const copied = await clickAfterRepublish('contextMenu.copy', snapshot => Object.assign(snapshot,
    { version: 2, workerSelection: range('body:t0:r0c0', 0, 2, 9) }));
  expect(copied.onError).not.toHaveBeenCalled();
  expect(copied.clipboard).toHaveBeenCalledWith(false);
  const cleared = await clickAfterRepublish('contextMenu.delete', snapshot => Object.assign(snapshot,
    { workerSelection: range('body:t0:r0c0', 0, 2, 9) }));
  expect(cleared.onError).not.toHaveBeenCalled();
  expect(cleared.submit.mock.calls[0]![0]).toEqual({ kind: 'replaceSelection',
    payload: { text: '', selection: range('body:t0:r0c0', 0, 2, 1) } });
});

test.each([
  ['select table, caret in another table', 'table.selectTable', { editState: inTable(1, 0), workerSelection: range('body:t1:r0c0', 0, 2, 1) }],
  ['delete row, caret in another row', 'table.deleteRow', { editState: inTable(0, 1), workerSelection: range('body:t0:r1c0', 0, 2, 1) }],
  ['delete table, after a document edit', 'table.deleteTable', { version: 2 }],
  ['delete, after a document edit', 'contextMenu.delete', { version: 2 }],
  ['delete, another range', 'contextMenu.delete', { workerSelection: range('body:t0:r0c0', 0, 1, 1) }],
  ['delete, reopened session', 'contextMenu.delete', { identity: { session: 's', incarnation: 1, epoch: 2 } }],
] as const)('context menu refuses a changed target: %s', async (_case, label, change) => {
  const result = await clickAfterRepublish(label, snapshot => Object.assign(snapshot, change));
  expect(result.execute).not.toHaveBeenCalled();
  expect(result.submit).not.toHaveBeenCalled();
  expect(result.clipboard).not.toHaveBeenCalled();
  expect(result.onError).toHaveBeenCalledWith(expect.objectContaining({ message: 'Context menu target changed' }));
});

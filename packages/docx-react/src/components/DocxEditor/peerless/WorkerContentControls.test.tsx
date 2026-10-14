import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { WorkerContentControls } from './WorkerContentControls';
import type { WorkerShellController } from './workerProperties';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

function fixture(frameEpoch = 2, readOnly = false) {
  const state = { status: 'ready', view: { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1,
    geometry: { frameEpoch: 2, layoutRevision: 3, coversThroughSeq: 1, pages: [{ page: { pageId: '1', pageIndex: 0 }, bounds: { x: 0, y: 0 } }] } } };
  const query = { version: 'native', value: { version: 'native', frameEpoch, layoutRevision: 3, controls: [0, 1].map(index => ({
    handle: `physical-${index}`, control: { alias: 'Flag', dataBound: false, effectiveLock: { content: false, control: false, known: true } },
    widget: { kind: 'checkbox', checked: false }, rects: [{ pageId: '1', pageIndex: 0, x: index * 100, y: 40, width: 12, height: 12 }],
  })) } };
  const applyQuery = mock(async () => ({ ok: true, status: 'executed' }));
  const controller = { subscribe: () => () => {}, getSnapshot: () => state, query: async () => query, applyQuery,
    environment: () => ({ readOnly, mode: 'editing', translate: () => 'Content control' }), ref: { focus() {} } } as unknown as WorkerShellController;
  return { controller, applyQuery };
}

test('widget writes carry the clicked physical handle and captured native token', async () => {
  const { controller, applyQuery } = fixture();
  const ui = render(<WorkerContentControls controller={controller} />);
  await waitFor(() => expect(ui.getAllByRole('checkbox')).toHaveLength(2));
  fireEvent.click(ui.getAllByRole('checkbox')[1]);
  await waitFor(() => expect(applyQuery).toHaveBeenCalledTimes(1));
  expect((applyQuery.mock.calls[0] as unknown[])[1]).toEqual({ kind: 'setContentControlValue', payload: {
    handle: 'physical-1', value: { kind: 'checkbox', checked: true }, expectVersion: 'native',
  } });
});

test('stale frame projections expose no widget targets', async () => {
  const { controller } = fixture(1);
  const ui = render(<WorkerContentControls controller={controller} />);
  await new Promise(resolve => setTimeout(resolve, 10));
  expect(ui.queryAllByRole('checkbox')).toHaveLength(0);
});

test('read-only widgets remain discoverable and cannot submit edits', async () => {
  const { controller, applyQuery } = fixture(2, true);
  const ui = render(<WorkerContentControls controller={controller} />);
  await waitFor(() => expect(ui.getAllByRole('checkbox')).toHaveLength(2));
  expect(ui.getAllByRole('checkbox').every(input => (input as HTMLInputElement).disabled)).toBe(true);
  fireEvent.click(ui.getAllByRole('checkbox')[1]);
  expect(applyQuery).not.toHaveBeenCalled();
});

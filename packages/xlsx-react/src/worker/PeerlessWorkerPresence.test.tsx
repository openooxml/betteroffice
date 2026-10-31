import { GlobalRegistrator } from '@happy-dom/global-registrator';
import type { WorkbookSheetSummary } from '@betteroffice/xlsx';
import type { AwarenessPeer, CollaborationAwareness } from '@betteroffice/xlsx/collaboration';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import { PeerlessWorkerPresence, xlsxPeerlessPresenceLoader } from './PeerlessWorkerSurface';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';
import { WorkerPresentationStore } from './WorkerPresentationStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
const originalLoad = xlsxPeerlessPresenceLoader.load;
type PresenceModule = Awaited<ReturnType<typeof originalLoad>>;

afterEach(() => { cleanup(); mock.restore(); });
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

function chunk() {
  let resolve!: (module: PresenceModule) => void;
  const promise = new Promise<PresenceModule>((done) => { resolve = done; });
  return { promise, resolve };
}

function presence() {
  const store = new WorkerPresentationStore();
  const grid = { startRow: 0, startCol: 0, rowOffsets: [0, 24, 48], colOffsets: [0, 96, 192] };
  store.setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
  store.setPainted({ source: 'worker', version: 'v1', sequence: 0, mergedRanges: [], geometry: grid,
    displayList: { width: 800, height: 600, commands: [], grid },
    request: { generation: 1, navigation: 0, sheet: 0, width: 800, height: 600, zoom: 1, dpr: 1,
      viewport: { x: 0, y: 0, width: 800, height: 600 } } });
  const run = { current: true, ready: true, failure: null, retiring: false } as unknown as PeerlessWorkbookSession;
  const sheets: WorkbookSheetSummary[] = [{ id: 'sheet:0', index: 0, name: 'Sheet 1' }];
  const peer: AwarenessPeer = { clientId: 202, clock: 1, lastSeen: Date.now(), cursorMovedAt: Date.now(),
    user: { name: 'Peer', color: '#0B57D0' },
    cursor: { sheet: 'sheet:0', anchor: { row: 0, col: 1 }, head: { row: 0, col: 1 } } };
  const listeners = new Set<(peers: readonly AwarenessPeer[]) => void>();
  const setCursor = mock<CollaborationAwareness['setCursor']>(() => {});
  const onAwareness = mock<CollaborationAwareness['onAwareness']>((listener) => {
    listeners.add(listener); listener([peer]);
    return () => { listeners.delete(listener); };
  });
  const provider = { peers: [peer], setCursor, onAwareness };
  return { props: { store, run, sheets, active: 0 }, provider, listeners, setCursor, onAwareness };
}

test('an editor without a collaboration provider never imports presence', () => {
  const load = spyOn(xlsxPeerlessPresenceLoader, 'load').mockImplementation(() => {
    throw new Error('Unexpected presence import');
  });
  const room = presence();
  const view = render(<PeerlessWorkerPresence {...room.props} provider={undefined} />);
  expect(view.container.childNodes.length).toBe(0);
  expect(load).not.toHaveBeenCalled();
  expect(room.listeners.size).toBe(0);
});

test('a delayed presence chunk subscribes, renders peers and selections, and cleans up on unmount', async () => {
  const module = await originalLoad();
  const pending = chunk();
  const load = spyOn(xlsxPeerlessPresenceLoader, 'load').mockReturnValue(pending.promise);
  const room = presence();
  const view = render(<PeerlessWorkerPresence {...room.props} provider={room.provider} />);
  expect(load).toHaveBeenCalledTimes(1);
  expect(room.onAwareness).not.toHaveBeenCalled();
  expect(view.queryByTestId('xlsx-presence-chip-202')).toBeNull();
  await act(async () => { pending.resolve(module); });
  await waitFor(() => expect(view.queryByTestId('xlsx-presence-chip-202')).not.toBeNull());
  expect(view.getByTestId('xlsx-remote-selection-202').style.width).toBe('96px');
  expect(room.setCursor).toHaveBeenLastCalledWith({ sheet: 'sheet:0', anchor: { row: 0, col: 0 }, head: { row: 0, col: 0 } });
  expect(room.listeners.size).toBe(1);
  await act(async () => { for (const listener of room.listeners) listener([]); });
  expect(view.queryByTestId('xlsx-presence-chip-202')).toBeNull();
  expect(view.queryByTestId('xlsx-remote-selection-202')).toBeNull();
  view.unmount();
  expect(room.listeners.size).toBe(0);
  expect(room.setCursor).toHaveBeenLastCalledWith(null);
});

test('a rejected presence chunk stays hidden and retries when the provider changes', async () => {
  const module = await originalLoad();
  const load = spyOn(xlsxPeerlessPresenceLoader, 'load').mockRejectedValue(new Error('Presence import failed'));
  const room = presence();
  const view = render(<PeerlessWorkerPresence {...room.props} provider={room.provider} />);
  await act(async () => {});
  expect(load).toHaveBeenCalledTimes(1);
  expect(view.container.childNodes.length).toBe(0);
  expect(room.onAwareness).not.toHaveBeenCalled();
  expect(room.setCursor).not.toHaveBeenCalled();
  load.mockResolvedValue(module);
  view.rerender(<PeerlessWorkerPresence {...room.props} provider={room.provider} />);
  await act(async () => {});
  expect(load).toHaveBeenCalledTimes(1);
  expect(view.container.childNodes.length).toBe(0);
  view.rerender(<PeerlessWorkerPresence {...room.props} provider={{ ...room.provider }} />);
  await waitFor(() => expect(view.queryByTestId('xlsx-presence-chip-202')).not.toBeNull());
  expect(load).toHaveBeenCalledTimes(2);
  expect(view.getByTestId('xlsx-remote-selection-202')).toBeTruthy();
  expect(room.listeners.size).toBe(1);
});

for (const retire of ['unmount', 'remove provider'] as const) {
  test(`${retire} during the presence import never attaches late awareness listeners`, async () => {
    const module = await originalLoad();
    const pending = chunk();
    spyOn(xlsxPeerlessPresenceLoader, 'load').mockReturnValue(pending.promise);
    const room = presence();
    const view = render(<PeerlessWorkerPresence {...room.props} provider={room.provider} />);
    expect(room.onAwareness).not.toHaveBeenCalled();
    if (retire === 'unmount') view.unmount();
    else view.rerender(<PeerlessWorkerPresence {...room.props} provider={undefined} />);
    await act(async () => { pending.resolve(module); });
    expect(room.listeners.size).toBe(0);
    expect(room.onAwareness).not.toHaveBeenCalled();
    expect(room.setCursor).not.toHaveBeenCalled();
    expect(view.queryByTestId('xlsx-presence-chip-202')).toBeNull();
  });
}

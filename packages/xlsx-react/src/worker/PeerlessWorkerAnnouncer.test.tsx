import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, it } from 'bun:test';
import type { DisplayList } from '@betteroffice/xlsx';
import { ActiveCellInputContext, PeerlessWorkerAnnouncer } from './PeerlessWorkerAnnouncer';
import { WorkerPresentationStore } from './WorkerPresentationStore';
import type { WorkerPaintResult } from './PeerlessWorkerPaintSource';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');

afterEach(async () => { await act(async () => { cleanup(); }); });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const frame = {
  request: { sheet: 0 },
  mergedRanges: [],
  displayList: {
    width: 300, height: 40,
    grid: { startRow: 0, startCol: 0, rowOffsets: [0, 20], colOffsets: [0, 64] },
    commands: [{ op: 'text', x: 2, y: 14, text: '42', fontSize: 11, color: '#000000', clip: { x: 0, y: 0, w: 64, h: 20 } }],
  } as DisplayList,
} as unknown as WorkerPaintResult;

const painted = (text: string) => ({
  ...frame,
  displayList: { ...frame.displayList, commands: text
    ? [{ op: 'text', x: 2, y: 14, text, fontSize: 11, color: '#000000', clip: { x: 0, y: 0, w: 64, h: 20 } }] : [] },
}) as unknown as WorkerPaintResult;

function setup() {
  const store = new WorkerPresentationStore();
  const focusTarget: { current: HTMLElement | null } = { current: null };
  const announceRef: { current: (() => void) | null } = { current: null };
  const Harness = ({ input }: { input: { at: string; value: string } | null }) =>
    <ActiveCellInputContext.Provider value={input}>
      <div ref={(element) => { focusTarget.current = element; }} tabIndex={0} data-testid="grid">
        <input data-testid="editor" />
        <PeerlessWorkerAnnouncer store={store} active={0} sheets={[]} focusTarget={focusTarget} announceRef={announceRef} />
      </div>
    </ActiveCellInputContext.Provider>;
  const view = render(<Harness input={null} />);
  const spoken = () => view.getByTestId('xlsx-live-region').textContent?.replace(/\u00a0$/, '');
  const select = () => act(async () => { store.setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } }); });
  return { store, view, Harness, spoken, select };
}

describe('peerless worker announcer', () => {
  it('corrects an announcement spoken before the cell data arrived', async () => {
    const { store, view, Harness, spoken, select } = setup();
    await act(async () => { view.getByTestId('grid').focus(); });
    await select();
    await waitFor(() => expect(spoken()).toBe('A1, empty'), { timeout: 3000 });
    await act(async () => { store.setPainted(frame); });
    view.rerender(<Harness input={{ at: '0:A1', value: '42' }} />);
    await waitFor(() => expect(spoken()).toBe('A1, 42'), { timeout: 3000 });
  });

  it('speaks a value that changes under the selected cell', async () => {
    const { store, view, Harness, spoken, select } = setup();
    await act(async () => { view.getByTestId('grid').focus(); store.setPainted(painted('42')); });
    view.rerender(<Harness input={{ at: '0:A1', value: '42' }} />);
    await select();
    await waitFor(() => expect(spoken()).toBe('A1, 42'), { timeout: 3000 });
    await act(async () => { store.setPainted(painted('')); });
    view.rerender(<Harness input={{ at: '0:A1', value: '' }} />);
    await waitFor(() => expect(spoken()).toBe('A1, empty'), { timeout: 3000 });
  });

  it('stays quiet about a value that changes while the cell editor has focus', async () => {
    const { store, view, Harness, spoken, select } = setup();
    await act(async () => { view.getByTestId('grid').focus(); store.setPainted(painted('42')); });
    view.rerender(<Harness input={{ at: '0:A1', value: '42' }} />);
    await select();
    await waitFor(() => expect(spoken()).toBe('A1, 42'), { timeout: 3000 });
    await act(async () => { view.getByTestId('editor').focus(); store.setPainted(painted('')); });
    view.rerender(<Harness input={{ at: '0:A1', value: '' }} />);
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 800)); });
    expect(spoken()).toBe('A1, 42');
  });
});

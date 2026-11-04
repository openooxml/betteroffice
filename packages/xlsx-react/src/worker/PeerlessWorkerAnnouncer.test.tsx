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
  const Harness = ({ input }: { input: { at: string; value: string; painted: object | null } | null }) =>
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
    view.rerender(<Harness input={{ at: '0:A1', value: '42', painted: frame }} />);
    await waitFor(() => expect(spoken()).toBe('A1, 42'), { timeout: 3000 });
  });

  it('speaks a value that changes under the selected cell, once its input matches the frame', async () => {
    const { store, view, Harness, spoken, select } = setup();
    const before = painted('=6*7 result');
    const after = painted('');
    await act(async () => { view.getByTestId('grid').focus(); store.setPainted(before); });
    view.rerender(<Harness input={{ at: '0:A1', value: '=6*7', painted: before }} />);
    await select();
    await waitFor(() => expect(spoken()).toBe('A1, =6*7 result, has formula'), { timeout: 3000 });
    await act(async () => { store.setPainted(after); });
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 400)); });
    expect(spoken()).toBe('A1, =6*7 result, has formula');
    view.rerender(<Harness input={{ at: '0:A1', value: '', painted: after }} />);
    await waitFor(() => expect(spoken()).toBe('A1, empty'), { timeout: 3000 });
  });

  it('stays quiet about a value that changes while the cell editor has focus', async () => {
    const { store, view, Harness, spoken, select } = setup();
    const before = painted('42');
    const after = painted('');
    await act(async () => { view.getByTestId('grid').focus(); store.setPainted(before); });
    view.rerender(<Harness input={{ at: '0:A1', value: '42', painted: before }} />);
    await select();
    await waitFor(() => expect(spoken()).toBe('A1, 42'), { timeout: 3000 });
    await act(async () => { view.getByTestId('editor').focus(); store.setPainted(after); });
    view.rerender(<Harness input={{ at: '0:A1', value: '', painted: after }} />);
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 800)); });
    expect(spoken()).toBe('A1, 42');
  });

  it('does not repeat a merged selection that scrolls out of the frame and back', async () => {
    const { store, view, Harness } = setup();
    const mergedRanges = [{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }];
    const shown = { ...painted('Head'), mergedRanges } as unknown as WorkerPaintResult;
    const away = { ...painted(''), mergedRanges: [], displayList: { width: 300, height: 40, commands: [],
      grid: { startRow: 50, startCol: 0, rowOffsets: [0, 20], colOffsets: [0, 64] } } } as unknown as WorkerPaintResult;
    const back = { ...shown } as WorkerPaintResult;
    const raw = () => view.getByTestId('xlsx-live-region').textContent;
    const wait = () => act(async () => { await new Promise((resolve) => setTimeout(resolve, 800)); });
    await act(async () => { view.getByTestId('grid').focus(); store.setPainted(shown); });
    view.rerender(<Harness input={{ at: '0:A1', value: 'Head', painted: shown }} />);
    await act(async () => { store.setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 1, col: 1 } }); });
    await waitFor(() => expect(raw()?.replace(/\u00a0$/, '')).toBe('A1, Head'), { timeout: 3000 });
    const first = raw();
    await act(async () => { store.setPainted(away); });
    view.rerender(<Harness input={{ at: '0:A1', value: 'Head', painted: away }} />);
    await wait();
    expect(raw()).toBe(first);
    await act(async () => { store.setPainted(back); });
    view.rerender(<Harness input={{ at: '0:A1', value: 'Head', painted: back }} />);
    await wait();
    expect(raw()).toBe(first);
  });
});

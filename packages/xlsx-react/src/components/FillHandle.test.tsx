import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, it, mock } from 'bun:test';
import type { CellAddr, CellRange } from '@betteroffice/xlsx';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
// testing-library must load after happy-dom installs its globals.
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
// the component under test imports React and expects the registered DOM.
const { FillHandle, FILL_HANDLE_SIZE, fillDestination, fillHandleAnchor, fillHandlePosition } =
  await import('./FillHandle');

const CELL_W = 96;
const CELL_H = 24;
const SOURCE: CellRange = { top: 0, left: 0, bottom: 0, right: 0 };
const SOURCE_RECT = { x: 0, y: 0, w: CELL_W, h: CELL_H };

const rectForRange = (range: CellRange) => ({
  x: range.left * CELL_W,
  y: range.top * CELL_H,
  w: (range.right - range.left + 1) * CELL_W,
  h: (range.bottom - range.top + 1) * CELL_H,
});
const cellAt = (clientX: number, clientY: number): CellAddr => ({
  row: Math.floor(clientY / CELL_H),
  col: Math.floor(clientX / CELL_W),
});

function mount(onFill = mock((_source: CellRange, _destination: CellRange) => {})) {
  const view = render(
    <FillHandle rect={SOURCE_RECT} source={SOURCE} rectForRange={rectForRange} cellAt={cellAt}
      onFill={onFill} ariaLabel="fill" />
  );
  return { view, onFill, handle: view.getByTestId('xlsx-fill-handle') };
}

const press = (handle: HTMLElement, x: number, y: number, pointerId = 1) =>
  fireEvent.pointerDown(handle, { button: 0, pointerId, clientX: x, clientY: y });
const move = (handle: HTMLElement, x: number, y: number, pointerId = 1) =>
  fireEvent.pointerMove(handle, { pointerId, clientX: x, clientY: y });
const release = (handle: HTMLElement, x: number, y: number, pointerId = 1) =>
  fireEvent.pointerUp(handle, { pointerId, clientX: x, clientY: y });

afterEach(async () => {
  await act(async () => { cleanup(); });
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

describe('fill geometry', () => {
  it('places the handle square on the selection bottom-right corner', () => {
    expect(fillHandlePosition({ x: 10, y: 20, w: 96, h: 24 })).toEqual({
      left: 10 + 96 - FILL_HANDLE_SIZE / 2,
      top: 20 + 24 - FILL_HANDLE_SIZE / 2,
      size: FILL_HANDLE_SIZE,
    });
  });

  it('fills the rectangle from the source top-left through the endpoint', () => {
    expect(fillDestination({ top: 1, left: 2, bottom: 2, right: 3 }, { row: 5, col: 4 }))
      .toEqual({ top: 1, left: 2, bottom: 5, right: 4 });
  });

  it('clamps an endpoint above or left of the source to its bottom-right corner', () => {
    expect(fillDestination({ top: 1, left: 2, bottom: 3, right: 4 }, { row: 0, col: 0 }))
      .toEqual({ top: 1, left: 2, bottom: 3, right: 4 });
  });

  it('anchors the handle on the true bottom-right cell when it is visible', () => {
    expect(fillHandleAnchor({ x: 0, y: 0, w: 96, h: 600 }, { x: 96, y: 220, w: 96, h: 24 }))
      .toEqual({ x: 96, y: 220, w: 96, h: 24 });
  });

  it('clamps the handle to the visible selection edge when the bottom-right is off-screen', () => {
    expect(fillHandleAnchor({ x: 0, y: 0, w: 96, h: 600 }, null))
      .toEqual({ x: 0, y: 0, w: 96, h: 600 });
    expect(fillHandleAnchor(null, null)).toBeNull();
  });
});

describe('FillHandle gesture', () => {
  it('renders the handle at the selection corner without a preview', () => {
    const { view, handle } = mount();
    expect(handle.style.left).toBe('92px');
    expect(handle.style.top).toBe('20px');
    expect(handle.style.width).toBe('8px');
    expect(handle.style.height).toBe('8px');
    expect(handle.style.pointerEvents).toBe('auto');
    expect(view.queryByTestId('xlsx-fill-preview')).toBeNull();
  });

  it('previews and commits a one-column fill dragged down', () => {
    const { view, handle, onFill } = mount();
    press(handle, 94, 22);
    move(handle, 150, 70);
    const preview = view.getByTestId('xlsx-fill-preview');
    expect(preview.style.left).toBe('0px');
    expect(preview.style.top).toBe('0px');
    expect(preview.style.width).toBe('192px');
    expect(preview.style.height).toBe('72px');
    release(handle, 150, 70);
    expect(onFill).toHaveBeenCalledTimes(1);
    expect(onFill).toHaveBeenCalledWith(SOURCE, { top: 0, left: 0, bottom: 2, right: 1 });
    expect(view.queryByTestId('xlsx-fill-preview')).toBeNull();
  });

  it('previews and commits a one-row fill dragged right', () => {
    const { view, handle, onFill } = mount();
    press(handle, 94, 22);
    move(handle, 200, 10);
    const preview = view.getByTestId('xlsx-fill-preview');
    expect(preview.style.width).toBe('288px');
    expect(preview.style.height).toBe('24px');
    release(handle, 200, 10);
    expect(onFill).toHaveBeenCalledWith(SOURCE, { top: 0, left: 0, bottom: 0, right: 2 });
  });

  it('does not fill or preview when dragged up and left', () => {
    const { view, handle, onFill } = mount();
    press(handle, 94, 22);
    move(handle, -20, -10);
    expect(view.queryByTestId('xlsx-fill-preview')).toBeNull();
    release(handle, -20, -10);
    expect(onFill).not.toHaveBeenCalled();
  });

  it('does not fill without a drag', () => {
    const { handle, onFill } = mount();
    press(handle, 94, 22);
    release(handle, 94, 22);
    expect(onFill).not.toHaveBeenCalled();
  });

  it('does not fill on a press and release over the neighbouring cell without moving', () => {
    const { handle, onFill } = mount();
    press(handle, 98, 26);
    release(handle, 98, 26);
    expect(onFill).not.toHaveBeenCalled();
  });

  it('cancels the preview and the pending fill on Escape', () => {
    const { view, handle, onFill } = mount();
    press(handle, 94, 22);
    move(handle, 150, 70);
    expect(view.getByTestId('xlsx-fill-preview')).toBeDefined();
    fireEvent.keyDown(window, { key: 'Escape' });
    expect(view.queryByTestId('xlsx-fill-preview')).toBeNull();
    release(handle, 150, 70);
    expect(onFill).not.toHaveBeenCalled();
  });

  it('uses the release point when the last move happened outside the handle', () => {
    const { handle, onFill } = mount();
    press(handle, 94, 22);
    release(handle, 150, 50);
    expect(onFill).toHaveBeenCalledWith(SOURCE, { top: 0, left: 0, bottom: 2, right: 1 });
  });
});

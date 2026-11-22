import { open } from './peerlessTestHarness';
import { expect, test } from 'bun:test';
import type { PptxAsyncEditorProps } from '../PptxEditor';
const { act, fireEvent, waitFor } = await import('@testing-library/react');

test('readOnlySelection={false} blocks selecting while slide navigation keeps working', async () => {
  const value = await open({ readOnly: true, readOnlySelection: false });
  act(() => {
    fireEvent.pointerDown(value.canvas, value.point());
    fireEvent.pointerUp(value.stage, value.point());
    fireEvent.pointerDown(value.canvas, { ...value.point(), detail: 2 });
    fireEvent.pointerUp(value.stage, value.point());
    fireEvent.contextMenu(value.canvas, value.point());
  });
  await act(async () => { await value.client.flush(); });
  expect(value.client.state.selection).toBeNull();
  expect(value.view.queryAllByTestId(/pptx-resize-/)).toHaveLength(0);
  expect(document.querySelector('[role="menu"]')).toBeNull();
  value.stage.focus();
  fireEvent.keyDown(value.stage, { key: 'PageDown' });
  await waitFor(() => expect(value.client.state.activeSlideId).toBe(value.client.state.slides[1]!.id));
});

test('read-only editors keep selection by default', async () => {
  const value = await open({ readOnly: true });
  act(() => {
    fireEvent.pointerDown(value.canvas, value.point());
    fireEvent.pointerUp(value.stage, value.point());
  });
  await waitFor(() => expect(value.client.state.selection).not.toBeNull());
});

test('turning readOnlySelection off resets the hover cursor', async () => {
  const options: Partial<PptxAsyncEditorProps> = { readOnly: true };
  const value = await open(options);
  act(() => { fireEvent.pointerMove(value.stage, value.point()); });
  expect(value.canvas.style.cursor).not.toBe('default');
  options.readOnlySelection = false;
  act(() => { value.view.rerender(value.ui()); });
  await waitFor(() => expect(value.canvas.style.cursor).toBe('default'));
});

import { expect, mock, test } from 'bun:test';
import { act, fireEvent } from '@testing-library/react';
import { resizeCursor } from '../interactions';
import { open } from './peerlessTestHarness';
import * as wasm from '../../../pptx/src/wasm/loader';

test('deleting every slide keeps document commands and history available', async () => {
  const saved = mock((_bytes: Uint8Array) => {});
  const value = await open({ onSave: saved });
  const access = await value.api.handleAsync();
  const slides = [...value.client.state.slides];
  for (const slide of slides) {
    await act(async () => { await access.deleteSlide(slide.id); });
  }
  expect(value.client.state.slides).toHaveLength(0);
  expect(value.view.getByTestId('pptx-worker-status').textContent).toBe('No slides');
  expect(value.view.queryByTestId('pptx-notes-textarea')).toBeNull();
  for (const command of ['undo', 'insertSlide', 'save'] as const) {
    expect(value.api.commands.getState(command).enabled).toBe(true);
  }
  for (const command of ['insertImage', 'exportPng', 'slideshow'] as const) {
    expect(value.api.commands.getState(command)).toMatchObject({ enabled: false, disabledReason: { code: 'slide-required' } });
  }
  await act(async () => { expect(await value.api.commands.execute('save', null)).toMatchObject({ ok: true }); });
  expect(saved).toHaveBeenCalledTimes(1);
  const empty = wasm.openPresentation(saved.mock.calls[0]![0]);
  try { expect(empty.snapshot().slides).toHaveLength(0); }
  finally { empty.dispose(); }
  await act(async () => { expect(await value.api.commands.execute('undo', null)).toMatchObject({ ok: true }); });
  expect(value.client.state.slides.map((slide) => slide.id)).toEqual([slides[1]!.id]);
  expect(value.api.commands.getState('redo').enabled).toBe(true);
  await act(async () => { expect(await value.api.commands.execute('redo', null)).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(0);
  await act(async () => { expect(await value.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(1);
  expect(value.client.state.activeSlideId).toBe(value.client.state.slides[0]!.id);
  expect(value.view.queryByTestId('pptx-worker-status')).toBeNull();
  const recovered = wasm.openPresentation(await value.api.saveAsync());
  try { expect(recovered.snapshot().slides).toHaveLength(1); }
  finally { recovered.dispose(); }
});

test('object_slide_history', async () => {
  const value = await open();
  await act(async () => {
    await value.api.commands!.execute('tool', { value: 'shape:rect' });
  });
  await act(async () => {
    fireEvent.pointerDown(value.canvas, value.point(320, 200));
    fireEvent.pointerUp(value.stage, value.point(320, 200));
    await value.api.flushPendingInput();
  });
  expect(value.client.state.selection?.kind).toBe('object');
  await act(async () => {
    expect(await value.api.commands!.execute('shapeFill', { color: '#abcdef' })).toMatchObject({ ok: true });
    expect(await value.api.commands!.execute('shapeStrokeWidth', { points: 3 })).toMatchObject({ ok: true });
    expect(await value.api.commands!.execute('zOrder', { value: 'back' })).toMatchObject({ ok: true });
  });
  const access = await value.api.handleAsync();
  const snapshot = await access.snapshot();
  const shape = snapshot.slides[0]!.shapes.find((shape) => shape.id === value.client.state.selection?.shapeId)!;
  expect(shape.resolvedFillColor?.toLowerCase()).toBe('#abcdef');
  await act(async () => { expect(await value.api.commands!.execute('insertSlide', { layoutPartPath: null })).toMatchObject({ ok: true }); });
  expect(value.client.state.slides).toHaveLength(3);
  expect(value.client.state.activeSlideId).toBe(value.client.state.slides[1]!.id);
  await act(async () => { await value.api.commands!.execute('undo', null); });
  expect(value.client.state.slides).toHaveLength(2);
  await act(async () => { await value.api.commands!.execute('redo', null); });
  expect(value.client.state.slides).toHaveLength(3);
  const inserted = value.client.state.activeSlideId!;
  await act(async () => { await access.moveSlide(inserted, 2); });
  expect(value.client.state.slides[2]!.id).toBe(inserted);
  await act(async () => { await access.deleteSlide(inserted); });
  expect(value.client.state.slides).toHaveLength(2);
  expect(value.client.state.activeSlideId).toBe(value.client.state.slides[1]!.id);
  await act(async () => { await access.removeShape(snapshot.slides[0]!.id, shape.id); });
  expect((await access.snapshot()).slides[0]!.shapes.some((candidate) => candidate.id === shape.id)).toBe(false);
  await act(async () => { await value.api.commands!.execute('undo', null); });
  expect((await access.snapshot()).slides[0]!.shapes.some((candidate) => candidate.id === shape.id)).toBe(true);
  const saved = await value.api.saveAsync();
  expect(saved).toBeInstanceOf(Uint8Array);
  await act(async () => { value.view.rerender(value.ui(true)); await value.api.flushPendingInput(); });
  await expect(access.deleteSlide(value.client.state.slides[0]!.id)).rejects.toThrow();
  expect(value.view.queryByRole('button', { name: 'Delete slide' })).toBeNull();
});

test('hover cursor follows the confirmed geometry under the pointer', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = { x: line.caretStops[2]!.x, y: line.y + line.height / 2 };
  const move = (x: number, y: number, buttons = 0, pointerId = 7) => {
    act(() => { fireEvent.pointerMove(value.stage, { clientX: x, clientY: y, pointerId, isPrimary: true, buttons }); });
    return value.canvas.style.cursor;
  };
  expect(value.canvas.style.cursor).toBe('default');
  expect(move(word.x, word.y)).toBe('text');
  expect(value.stage.style.cursor).toBe('');
  expect(move(value.box.x + 1, value.box.y + value.box.h / 2)).toBe('move');
  expect(move(value.list.width - 2, value.list.height - 2)).toBe('default');
  expect(move(word.x, word.y, 1)).toBe('default');
  const last = value.box.lines[value.box.lines.length - 1]!;
  expect(move(word.x, (last.y + last.height + value.box.y + value.box.h) / 2)).toBe('move');
  act(() => { fireEvent.pointerDown(value.canvas, value.point()); });
  expect(move(word.x, word.y, 0, 8)).toBe('move');
  act(() => { fireEvent.pointerUp(value.stage, value.point()); });
  expect(move(value.list.width - 2, value.list.height - 2)).toBe('default');
  await act(async () => { await value.api.commands!.execute('tool', { value: 'shape:rect' }); });
  expect(value.canvas.style.cursor).toBe('crosshair');
  expect(move(word.x, word.y)).toBe('crosshair');
});

test('the capturing stage shows the press cursor until the press ends', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...word, buttons: 1 }); });
  expect(value.stage.setPointerCapture).toHaveBeenCalled();
  expect(value.stage.style.cursor).toBe('text');
  act(() => { fireEvent.pointerUp(value.stage, word); });
  expect(value.stage.style.cursor).toBe('');
  expect(value.canvas.style.cursor).toBe('text');
  await value.select();
  const corner = value.point(value.box.x + value.box.w, value.box.y + value.box.h);
  act(() => { fireEvent.pointerDown(value.view.getByTestId('pptx-resize-se'), { ...corner, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe(resizeCursor('se'));
  act(() => { fireEvent.pointerCancel(value.stage, corner); });
  expect(value.stage.style.cursor).toBe('');
  await act(async () => { await value.api.commands!.execute('tool', { value: 'shape:rect' }); });
  act(() => { fireEvent.pointerDown(value.view.getByTestId('pptx-resize-se'), { ...corner, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe('crosshair');
  act(() => { fireEvent.lostPointerCapture(value.stage, corner); });
  expect(value.stage.style.cursor).toBe('');
});

test('a press resolves its cursor at the press point', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerMove(value.canvas, { ...word, buttons: 0 }); });
  expect(value.canvas.style.cursor).toBe('text');
  const body = value.point(value.box.x + 1, value.box.y + value.box.h / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...body, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe('move');
  act(() => { fireEvent.pointerUp(value.stage, body); });
});

test('right-clicking text leaves the stage cursor unset', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...word, button: 2, buttons: 2 }); });
  expect(value.stage.style.cursor).toBe('');
});

test('pressing empty slide space leaves the stage cursor unset', async () => {
  const value = await open();
  const empty = value.point(value.list.width - 2, value.list.height - 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...empty, buttons: 1 }); });
  expect(value.stage.setPointerCapture).not.toHaveBeenCalled();
  expect(value.stage.style.cursor).toBe('');
});

test('another pointer leaves the captured press cursor in place', async () => {
  const value = await open();
  const line = value.box.lines[0]!;
  const word = value.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(value.canvas, { ...word, buttons: 1 }); });
  expect(value.stage.hasPointerCapture(7)).toBe(true);
  expect(value.stage.style.cursor).toBe('text');
  const second = { ...value.point(value.list.width - 2, value.list.height - 2), pointerId: 8, isPrimary: false };
  act(() => { fireEvent.pointerDown(value.stage, { ...second, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe('text');
  act(() => { fireEvent.pointerUp(value.stage, { ...second, buttons: 0 }); });
  expect(value.stage.style.cursor).toBe('text');
  act(() => { fireEvent.pointerUp(value.stage, { ...word, buttons: 0 }); });
  expect(value.stage.style.cursor).toBe('');
});

test('a resize handle keeps the resize cursor off the canvas', async () => {
  const value = await open();
  await value.select();
  const corner = value.point(value.box.x + value.box.w, value.box.y + value.box.h);
  const handle = value.view.getByTestId('pptx-resize-se');
  act(() => { fireEvent.pointerMove(handle, { ...corner, buttons: 0 }); });
  expect(value.canvas.style.cursor).not.toContain('resize');
  act(() => { fireEvent.pointerDown(handle, { ...corner, buttons: 1 }); });
  expect(value.stage.style.cursor).toBe(resizeCursor('se'));
  act(() => { fireEvent.pointerCancel(value.stage, { ...corner, buttons: 0 }); });
  expect(value.canvas.style.cursor).not.toContain('resize');
});

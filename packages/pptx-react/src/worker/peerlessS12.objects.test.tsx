import { expect, mock, test } from 'bun:test';
import { act, fireEvent } from '@testing-library/react';
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

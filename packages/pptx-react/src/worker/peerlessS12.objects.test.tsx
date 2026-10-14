import { expect, test } from 'bun:test';
import { act, fireEvent } from '@testing-library/react';
import { open } from './peerlessTestHarness';

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

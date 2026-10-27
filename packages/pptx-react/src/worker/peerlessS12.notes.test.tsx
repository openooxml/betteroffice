import { expect, mock, test } from 'bun:test';
import type { DeckSnapshot, PresentationHandle } from '@betteroffice/pptx';
import { open, replyControl, story, type } from './peerlessTestHarness';
import * as wasm from '../../../pptx/src/wasm/loader';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

type NotesAccess = { [K in 'setSlideNotes' | 'addComment' | 'replyToComment' | 'setCommentStatus' | 'setCommentPosition' | 'removeComment' | 'setCommentFlavor' | 'comments']:
  (...args: Parameters<PresentationHandle[K]>) => Promise<ReturnType<PresentationHandle[K]>> };

test('notes drafts queue behind typing in input order', async () => {
  const error = mock(() => {});
  const value = await open({ onError: error });
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  replyControl.hold();
  try {
    act(() => { type(value.view.getByLabelText('Slide text'), 'Typed'); });
    await waitFor(() => expect(replyControl.count()).toBe(1));
    act(() => {
      const notes = value.view.getByTestId('pptx-notes-textarea');
      fireEvent.change(notes, { target: { value: 'A' } });
      fireEvent.change(notes, { target: { value: 'AB' } });
    });
    act(() => { type(value.view.getByLabelText('Slide text'), 'NEXT'); });
  } finally { replyControl.release(); }
  await act(async () => { await value.api.flushPendingInput(); });
  await waitFor(() => expect((value.view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement).value).toBe('AB'));
  expect((await (await value.api.handleAsync()).snapshot()).slides[0]!.notes).toBe('AB');
  expect((await story(value.api)).text).toBe(`TypedNEXT${target.text}`);
  expect(error).not.toHaveBeenCalled();
});

test('notes typed before a slide switch is acknowledged stay on their slide', async () => {
  const error = mock(() => {});
  const value = await open({ onError: error });
  const second = value.client.state.slides[1]!.id;
  const notes = () => value.view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement;
  replyControl.hold();
  try {
    act(() => { void value.api.goToSlide(2); });
    await waitFor(() => expect(replyControl.count()).toBe(1));
    act(() => { fireEvent.change(notes(), { target: { value: 'First' } }); });
    act(() => { replyControl.releaseNext(); });
    await waitFor(() => expect(value.client.state.activeSlideId).toBe(second));
    act(() => { fireEvent.change(notes(), { target: { value: `${notes().value}x` } }); });
  } finally { replyControl.release(); }
  await act(async () => { await value.api.flushPendingInput(); });
  expect((await (await value.api.handleAsync()).snapshot()).slides.map((slide) => slide.notes ?? '')).toEqual(['First', 'x']);
  await waitFor(() => expect(notes().value).toBe('x'));
  expect(error).not.toHaveBeenCalled();
});

for (const method of ['setSlideNotes', 'addComment'] as const) {
  test(`typing queues behind an unacknowledged ${method}`, async () => {
    const error = mock(() => {});
    const value = await open({ onError: error });
    const target = await story(value.api);
    await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
    const access = await value.api.handleAsync();
    replyControl.hold();
    const invoked = method === 'setSlideNotes' ? access.setSlideNotes(target.slideId, 'Notes')
      : access.addComment(target.slideId, { author: 'Writer', text: 'Comment', created: '2026-01-01T00:00:00Z' });
    try {
      await waitFor(() => expect(replyControl.count()).toBe(1));
      act(() => { type(value.view.getByLabelText('Slide text'), 'NEXT'); });
    } finally { replyControl.release(); }
    await act(async () => { await invoked; await value.api.flushPendingInput(); });
    expect((await story(value.api)).text).toBe(`NEXT${target.text}`);
    if (method === 'setSlideNotes') expect((await access.snapshot()).slides[0]!.notes).toBe('Notes');
    else expect(await access.comments()).toMatchObject([{ text: 'Comment' }]);
    expect(error).not.toHaveBeenCalled();
  });
}

test('typing survives an unacknowledged refused comment', async () => {
  const value = await open();
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  const access = await value.api.handleAsync();
  replyControl.hold();
  const refused = access.addComment(target.slideId, { author: 'Writer', text: 'Comment', created: 'invalid' })
    .catch((error: unknown) => error);
  try {
    await waitFor(() => expect(replyControl.count()).toBe(1));
    act(() => { type(value.view.getByLabelText('Slide text'), 'NEXT'); });
  } finally { replyControl.release(); }
  await act(async () => {
    expect(await refused).toMatchObject({ code: 'invalid-arguments' });
    await value.api.flushPendingInput();
  });
  expect((await story(value.api)).text).toBe(`NEXT${target.text}`);
  expect(await access.comments()).toEqual([]);
  expect(value.api.failure).toBeUndefined();
});

test('invalid comments return a typed payload refusal and preserve editing', async () => {
  const value = await open();
  const access = await value.api.handleAsync();
  const target = await story(value.api);
  await act(async () => {
    await expect(access.addComment(target.slideId, { author: 'Writer', text: '', created: '2026-01-01T00:00:00Z' }))
      .rejects.toMatchObject({ code: 'invalid-arguments' });
    await access.addComment(target.slideId, { author: 'Writer', text: 'Valid', created: '2026-01-01T00:00:00Z' });
    await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 });
    type(value.view.getByLabelText('Slide text'), 'NEXT');
    await value.api.flushPendingInput();
  });
  expect(value.api.failure).toBeUndefined();
  expect((await story(value.api)).text).toBe(`NEXT${target.text}`);
  expect(await access.comments()).toMatchObject([{ text: 'Valid' }]);
});

test('notes_comments_change', async () => {
  const changed = mock((_snapshot: DeckSnapshot) => {});
  const value = await open({ onChange: changed });
  const notes = value.view.getByTestId('pptx-notes-textarea');
  await act(async () => { fireEvent.change(notes, { target: { value: 'Speaker notes' } }); await value.api.flushPendingInput(); });
  const access = await value.api.handleAsync() as unknown as NotesAccess;
  expect((await (await value.api.handleAsync()).snapshot()).slides[0]!.notes).toBe('Speaker notes');
  await act(async () => {
    await access.setCommentFlavor('modern');
    const root = await access.addComment(value.client.state.activeSlideId!, { author: 'Writer', text: 'Discuss', created: '2026-01-01T00:00:00Z' });
    await access.replyToComment(root.commentId, { author: 'Reader', text: 'Agreed', created: '2026-01-01T00:00:01Z' });
    await access.setCommentPosition(root.commentId, { xEmu: 100, yEmu: 200 });
    await access.setCommentStatus(root.commentId, true);
  });
  const comments = await access.comments();
  expect(comments).toHaveLength(2);
  expect(comments.find((comment) => !comment.parentId)).toMatchObject({ resolved: true, xEmu: 100, yEmu: 200 });
  await waitFor(() => expect(changed.mock.calls[changed.mock.calls.length - 1]?.[0]).toMatchObject({ commentFlavor: 'modern', comments }));
  const saved = wasm.openPresentation(await value.api.saveAsync());
  try {
    expect(saved.snapshot().slides[0]!.notes).toBe('Speaker notes');
    const normalized = (values: typeof comments) => values.map(({ id, parentId, ...comment }) => ({
      ...comment, parent: parentId ? values.findIndex((value) => value.id === parentId) : null,
    }));
    expect(normalized(saved.comments())).toEqual(normalized(comments));
  } finally { saved.dispose(); }
  await act(async () => { value.view.rerender(value.ui(true)); await value.api.flushPendingInput(); });
  expect((value.view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement).disabled).toBe(true);
  await expect(access.setSlideNotes(value.client.state.activeSlideId!, 'Blocked')).rejects.toThrow();
});

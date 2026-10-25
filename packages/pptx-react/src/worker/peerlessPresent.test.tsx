import { open, story } from './peerlessTestHarness';
import { expect, spyOn, test } from 'bun:test';
import { decodeCommandDisplayList } from '@betteroffice/pptx/worker';
import type { PresentationCommandFrame } from '@betteroffice/pptx/worker';
import { EditorToolbar } from '../components/EditorToolbar';
const { act, fireEvent, waitFor } = await import('@testing-library/react');

function counter(dialog: HTMLElement): string | null {
  return dialog.querySelector('[aria-live="polite"]')?.textContent ?? null;
}

function texts(frame: PresentationCommandFrame): string {
  return decodeCommandDisplayList(frame).primitives
    .flatMap((primitive) => primitive.kind === 'textBox' ? primitive.paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)) : [])
    .join('');
}

test('the default editor presents from the current slide', async () => {
  const value = await open({ initialSlide: 2 });
  const present = value.view.getByTestId('pptx-present');
  expect(present.textContent).toContain('Present');
  await act(async () => { fireEvent.click(present); });
  const dialog = await waitFor(() => value.view.getByRole('dialog'));
  expect(dialog.getAttribute('aria-label')).toBe('Slide presentation');
  expect(counter(dialog)).toBe('2 / 2');
});

test('presenting advances with keys and clicks and exits on Escape', async () => {
  const value = await open();
  await act(async () => { fireEvent.click(value.view.getByTestId('pptx-present')); });
  const dialog = await waitFor(() => value.view.getByRole('dialog'));
  expect(counter(dialog)).toBe('1 / 2');
  act(() => { fireEvent.keyDown(dialog, { key: 'ArrowRight' }); });
  expect(counter(dialog)).toBe('2 / 2');
  act(() => { fireEvent.keyDown(dialog, { key: 'ArrowLeft' }); });
  expect(counter(dialog)).toBe('1 / 2');
  act(() => { fireEvent.click(dialog); });
  expect(counter(dialog)).toBe('2 / 2');
  act(() => { fireEvent.keyDown(dialog, { key: 'Escape' }); });
  await waitFor(() => expect(value.view.queryByRole('dialog')).toBeNull());
  expect(value.view.getByTestId('pptx-present')).toBeDefined();
});

test('an edit made while presenting shows on the presented slide', async () => {
  const value = await open();
  const text = await story(value.api);
  const target = { kind: 'range' as const, slideId: text.slideId, shapeId: text.shapeId, storyId: text.storyId, start: 0, end: 0 };
  await act(async () => {
    expect(await value.api.applyEdits({ expectVersion: await value.api.version(), steps: [{ op: 'replaceText', target, text: 'First ' }] }))
      .toMatchObject({ ok: true, applied: true });
  });
  const read = value.client.thumbnail.bind(value.client);
  const painted: string[] = [];
  const thumbnail = spyOn(value.client, 'thumbnail').mockImplementation(async (...args: Parameters<typeof read>) => {
    const reply = await read(...args);
    if (args[0] === text.slideId && args.length === 1 && reply.status === 'frame') painted.push(texts(reply.frame));
    return reply;
  });
  try {
    await act(async () => { fireEvent.click(value.view.getByTestId('pptx-present')); });
    await waitFor(() => expect(painted.some((value) => value.includes('First '))).toBe(true));
    await act(async () => {
      expect(await value.api.applyEdits({ expectVersion: await value.api.version(), steps: [{ op: 'replaceText', target, text: 'Second ' }] }))
        .toMatchObject({ ok: true, applied: true });
    });
    await waitFor(() => expect(painted[painted.length - 1]).toContain('Second First '));
    expect(counter(value.view.getByRole('dialog'))).toBe('1 / 2');
  } finally { thumbnail.mockRestore(); }
});

test('a read-only viewer without a toolbar still shows Present', async () => {
  const value = await open({ readOnly: true });
  expect(value.view.queryByTestId('pptx-editor-toolbar')).toBeNull();
  const present = value.view.getByTestId('pptx-present');
  await act(async () => { fireEvent.click(present); });
  await waitFor(() => expect(value.view.getByRole('dialog')).toBeDefined());
});

test('a hidden toolbar hides the Present button', async () => {
  const value = await open({ toolbar: null });
  expect(value.view.queryByTestId('pptx-present')).toBeNull();
  await act(async () => { value.view.rerender(value.ui(true)); });
  expect(value.view.queryByTestId('pptx-present')).toBeNull();
});

test('a custom toolbar replaces the Present button', async () => {
  const value = await open({ toolbar: <EditorToolbar mode="commands"><span data-testid="custom-toolbar" /></EditorToolbar> });
  expect(value.view.getByTestId('custom-toolbar')).toBeDefined();
  expect(value.view.queryByTestId('pptx-present')).toBeNull();
});

import { expect, test } from 'bun:test';
import { act, fireEvent, waitFor } from '@testing-library/react';
import { open } from './peerlessTestHarness';

const mirror = (view: Awaited<ReturnType<typeof open>>['view']) => view.getByTestId('pptx-slide-mirror');
const live = (view: Awaited<ReturnType<typeof open>>['view']) => view.getByTestId('pptx-live-region');

test('the slide mirror is a labelled document holding the text as whole paragraphs', async () => {
  const { view } = await open();
  const document = mirror(view);
  expect(document.getAttribute('role')).toBe('document');
  expect(document.getAttribute('aria-label')).toBe('Slide 1 of 2');
  expect([...document.querySelectorAll('p')].map((node) => node.textContent)).toEqual(['alpha beta', 'x', 'alpha beta']);
  expect(document.closest('[role="application"]')).not.toBeNull();
  expect(view.getAllByRole('application')).toHaveLength(1);
});

test('slide thumbnails carry the slide position in their names', async () => {
  const { view } = await open();
  expect(view.getByRole('button', { name: 'Slide 2 of 2' })).toBeTruthy();
});

test('moving to another slide is announced, and only while focus is in the editor', async () => {
  const { view, stage } = await open();
  act(() => { (stage as HTMLElement).focus(); });
  fireEvent.keyDown(stage, { key: 'PageDown' });
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 2 of 2'));
  await waitFor(() => expect(live(view).textContent?.trim()).toBe('Slide 2 of 2'));
  act(() => { (stage as HTMLElement).blur(); });
  const before = live(view).textContent;
  fireEvent.click(view.getByRole('button', { name: 'Slide 1 of 2' }));
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 1 of 2'));
  act(() => { (document.activeElement as HTMLElement | null)?.blur(); });
  await new Promise((resolve) => setTimeout(resolve, 250));
  expect(live(view).textContent).toBe(before);
});

test('selecting a shape is announced by name', async () => {
  const { view, stage, select } = await open();
  act(() => { (stage as HTMLElement).focus(); });
  await select();
  await waitFor(() => expect(live(view).textContent?.trim()).toBe('Text selected'));
});

test('page keys in the mirror change slides, and other keys leave the selection alone', async () => {
  const { view, select, client } = await open();
  await select();
  const document = mirror(view);
  act(() => { document.focus(); });
  fireEvent.keyDown(document, { key: 'Delete' });
  await new Promise((resolve) => setTimeout(resolve, 50));
  expect(client.state.selection).toMatchObject({ kind: 'object' });
  fireEvent.keyDown(document, { key: 'PageDown' });
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 2 of 2'));
});

test('a host page document around the editor leaves the stage keys working', async () => {
  const { view, stage } = await open();
  view.container.setAttribute('role', 'document');
  act(() => { (stage as HTMLElement).focus(); });
  fireEvent.keyDown(stage, { key: 'PageDown' });
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 2 of 2'));
});

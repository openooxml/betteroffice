import { open, story } from './peerlessTestHarness';
import { expect, mock, test } from 'bun:test';
const { act, fireEvent, waitFor } = await import('@testing-library/react');

function compose(input: HTMLElement, type: string, data = '') {
  const event = new Event(type, { bubbles: true });
  Object.defineProperty(event, 'data', { value: data });
  fireEvent(input, event);
}

test('scaled_localized_preedit', async () => {
  const value = await open({ i18n: { editor: { textInput: 'Folientext' } } });
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  await act(async () => { expect(await value.api.commands!.execute('zoom', { scale: 2 })).toMatchObject({ ok: true }); });
  await act(async () => { await value.api.flushPendingInput(); });
  const input = value.view.getByLabelText('Folientext');
  await act(async () => {
    input.focus();
    compose(input, 'compositionstart');
    await value.api.flushPendingInput();
  });
  await act(async () => { compose(input, 'compositionupdate', '日本'); });
  await waitFor(() => expect(value.view.queryByTestId('pptx-peerless-preedit')).not.toBeNull());
  const preedit = value.view.getByTestId('pptx-peerless-preedit');
  expect(preedit.textContent).toBe('日本');
  expect(preedit.style.left).toBe(input.style.left);
  expect(parseFloat(input.style.height)).toBeGreaterThan(1);
  await act(async () => { compose(input, 'compositionend', '日本'); await value.api.flushPendingInput(); });
  expect((await story(value.api)).text.startsWith('日本')).toBe(true);
  expect(value.view.queryByTestId('pptx-peerless-preedit')).toBeNull();
});

test('copy_cut_waits_for_clipboard_and_preserves_newer_selection', async () => {
  const value = await open();
  const target = await story(value.api);
  const select = async (start: number, end: number) => act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start, end }); });
  const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  let release!: () => void;
  const write = mock(() => Promise.resolve());
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: write } });
  try {
    await select(0, 5);
    const input = value.view.getByLabelText('Slide text');
    await act(async () => { fireEvent.keyDown(input, { key: 'c', ctrlKey: true }); });
    await waitFor(() => expect(write).toHaveBeenCalledWith('alpha'));
    write.mockImplementation(() => new Promise<void>((resolve) => { release = resolve; }));
    await act(async () => { fireEvent.keyDown(input, { key: 'x', ctrlKey: true }); });
    await waitFor(() => expect(write).toHaveBeenCalledTimes(2));
    await select(6, 10);
    await act(async () => { release(); await value.api.flushPendingInput(); });
    expect((await story(value.api)).text).toBe(target.text);
    write.mockImplementation(() => Promise.resolve());
    await act(async () => { fireEvent.keyDown(input, { key: 'x', ctrlKey: true }); });
    await waitFor(async () => expect((await story(value.api)).text.startsWith('alpha ')).toBe(true));
    await waitFor(async () => expect((await story(value.api)).text.startsWith('alpha beta')).toBe(false));
  } finally {
    if (original) Object.defineProperty(navigator, 'clipboard', original);
    else Reflect.deleteProperty(navigator, 'clipboard');
  }
});

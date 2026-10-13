import { open, story, type } from './peerlessTestHarness';
import { expect, test } from 'bun:test';
import { definePptxPlugin } from '../plugins/definePptxPlugin';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

test('slide text clicks and API focus keep typing out of plugin textareas', async () => {
  const plugin = definePptxPlugin({ id: 'textarea', createState: () => null,
    overlay: () => <textarea data-testid="plugin-textarea" style={{ pointerEvents: 'auto' }} /> });
  const value = await open({ plugins: [plugin] });
  await waitFor(() => expect(value.view.queryByTestId('plugin-textarea')).toBeTruthy());
  const input = value.view.getByRole('textbox', { name: 'Slide text' });
  const before = await story(value.api);
  const line = value.box.lines[0]!;
  const point = value.point(value.box.x + 40, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(value.canvas, point); fireEvent.pointerUp(value.stage, point); });
  await value.api.flushPendingInput();
  expect(value.client.state.selection?.kind).not.toBe('object');
  expect(document.activeElement === input).toBe(true);
  await act(async () => { type(document.activeElement as HTMLElement, 'Z'); await value.api.flushPendingInput(); });
  expect((await story(value.api)).text).not.toBe(before.text);
  expect((value.view.getByTestId('plugin-textarea') as HTMLTextAreaElement).value).toBe('');
  act(() => { value.view.getByTestId('plugin-textarea').focus(); value.api.focus(); });
  expect(document.activeElement === input).toBe(true);
});

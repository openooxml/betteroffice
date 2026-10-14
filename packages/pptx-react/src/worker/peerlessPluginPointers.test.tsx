import { open } from './peerlessTestHarness';
import { expect, test } from 'bun:test';
import { createPortal } from 'react-dom';
import { definePptxPlugin } from '../plugins/definePptxPlugin';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

test('plugin overlay portals keep pointer input out of the insertion tool', async () => {
  let clicks = 0;
  const plugin = definePptxPlugin({ id: 'portal', createState: () => null,
    overlay: () => createPortal(<input data-testid="plugin-portal" onPointerDown={() => { clicks++; }} />, document.body) });
  const value = await open({ plugins: [plugin] });
  await waitFor(() => expect(document.querySelector('[data-testid="plugin-portal"]')).toBeTruthy());
  const portal = document.querySelector<HTMLInputElement>('[data-testid="plugin-portal"]')!;
  await act(async () => { expect(await value.api.commands!.execute('tool', { value: 'textBox' })).toMatchObject({ ok: true }); });
  const version = await value.api.version();
  await act(async () => {
    fireEvent.pointerDown(portal, value.point(300, 200));
    fireEvent.pointerMove(portal, { ...value.point(380, 260), buttons: 1 });
    fireEvent.pointerUp(portal, value.point(380, 260));
    await value.api.flushPendingInput();
  });
  expect(clicks).toBe(1);
  expect(value.commands.filter((command) => command.kind === 'insertTextBox')).toHaveLength(0);
  expect(await value.api.version()).toBe(version);
  await act(async () => {
    fireEvent.pointerDown(value.canvas, value.point(300, 200));
    fireEvent.pointerUp(value.stage, value.point(380, 260));
    await value.api.flushPendingInput();
  });
  expect(value.commands.filter((command) => command.kind === 'insertTextBox')).toHaveLength(1);
});

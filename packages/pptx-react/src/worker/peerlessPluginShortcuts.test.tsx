import { open } from './peerlessTestHarness';
import { expect, test } from 'bun:test';
import { definePptxPlugin } from '../plugins/definePptxPlugin';

const { act, fireEvent } = await import('@testing-library/react');

test('contributed shortcuts run from slide input and leave plugin fields alone', async () => {
  let marks = 0;
  const plugin = definePptxPlugin({ id: 'shortcut', createState: () => null,
    panel: { title: 'Shortcut', placement: 'right', render: () => <input data-testid="plugin-field" /> },
    commands: [{ id: 'mark', label: 'Mark', mutatesDocument: false, shortcuts: ['Mod+Shift+K'],
      execute() { marks++; return { ok: true, status: 'executed' }; } }],
  });
  const value = await open({ plugins: [plugin] });
  const input = value.view.getByRole('textbox', { name: 'Slide text' });
  const chord = { key: 'k', ctrlKey: true, shiftKey: true };
  await act(async () => { input.focus(); fireEvent.keyDown(input, chord); });
  expect(marks).toBe(1);
  await act(async () => {
    fireEvent.keyDown(input, { ...chord, repeat: true });
    fireEvent.keyDown(input, { ...chord, isComposing: true });
    const field = value.view.getByTestId('plugin-field');
    field.focus(); fireEvent.keyDown(field, chord);
  });
  expect(marks).toBe(1);
});

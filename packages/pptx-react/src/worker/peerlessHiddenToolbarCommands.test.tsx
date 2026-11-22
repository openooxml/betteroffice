import { open } from './peerlessTestHarness';
import { expect, test } from 'bun:test';

test('hiddenToolbarCommands leaves the listed controls out of the default toolbar', async () => {
  const value = await open({ hiddenToolbarCommands: ['slideshow', 'proposalsPanel', 'bold', 'undo', 'redo'] });
  const toolbar = value.view.getByTestId('pptx-formatting-toolbar');
  expect(value.view.queryByTestId('pptx-present')).toBeNull();
  expect(value.view.queryByTestId('pptx-proposals-button')).toBeNull();
  expect(toolbar.querySelector('[aria-label="Bold"]')).toBeNull();
  expect(toolbar.querySelector('[role="group"][aria-label="History"]')).toBeNull();
  expect(toolbar.querySelector('[aria-label="Italic"]')).not.toBeNull();
  expect(toolbar.querySelector('[role="group"][aria-label="Text formatting"]')).not.toBeNull();
});

import { open, story } from './peerlessTestHarness';
import { expect, test } from 'bun:test';
const { act, waitFor, within } = await import('@testing-library/react');

test('the proposals panel docks beside the slide with the legacy panel styling', async () => {
  const value = await open();
  const target = await story(value.api);
  const access = await value.api.handleAsync();
  await act(async () => {
    await access.propose('reviewer', 'Shorter title', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Short' }]);
    await value.api.commands.execute('proposalsPanel', { open: true });
  });
  const panel = await waitFor(() => value.view.getByTestId('pptx-proposals-panel'));
  expect(value.stage.contains(panel)).toBe(false);
  expect(panel.parentElement).toBe(value.stage.parentElement);
  expect(panel.style.width).toBe('340px');
  expect(panel.style.borderLeft).toContain('solid');
  expect(panel.style.overflowY).toBe('auto');
  const card = within(panel).getByTestId('pptx-proposal');
  expect(card.style.borderRadius).toBe('8px');
  expect(card.style.borderLeft).toContain('solid');
  expect(within(card).getByTestId('pptx-proposal-accept').style.background).not.toBe('');
  expect(within(card).getByTestId('pptx-proposal-reject').style.borderRadius).toBe('5px');
  expect(within(card).getByTestId('pptx-proposal-preview').style.borderRadius).toBe('5px');
});

test('speaker notes and slide actions are styled like the legacy editor', async () => {
  const value = await open();
  const notes = value.view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement;
  expect(notes.style.borderRadius).toBe('6px');
  const label = value.view.getByText('Speaker notes');
  expect(label.style.textTransform).toBe('uppercase');
  for (const name of ['Add slide', 'Delete slide']) {
    const button = value.view.getByRole('button', { name });
    expect(button.style.borderRadius).not.toBe('');
    expect(button.style.background).not.toBe('');
  }
});

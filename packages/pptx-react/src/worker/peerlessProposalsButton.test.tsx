import { open, story } from './peerlessTestHarness';
import { expect, spyOn, test } from 'bun:test';
import { PptxDocument, PptxRenderer } from '../../../pptx/src/wasm/engine';
const { act, fireEvent, waitFor } = await import('@testing-library/react');

test('the default editor shows agent proposals and opens their panel', async () => {
  const value = await open();
  const button = value.view.getByTestId('pptx-proposals-button');
  expect(value.view.getByTestId('pptx-proposals-count').textContent).toBe('0');
  expect(button.getAttribute('aria-expanded')).toBe('false');
  const target = await story(value.api);
  const access = await value.api.handleAsync();
  await act(async () => {
    await access.propose('reviewer', 'Shorter title', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Short' }]);
    await value.api.flushPendingInput();
  });
  await waitFor(() => expect(value.view.getByTestId('pptx-proposals-count').textContent).toBe('1'));
  await act(async () => { fireEvent.click(button); });
  const panel = await waitFor(() => value.view.getByTestId('pptx-proposals-panel'));
  expect(panel.textContent).toContain('Shorter title');
  expect(button.getAttribute('aria-expanded')).toBe('true');
  await act(async () => { fireEvent.click(button); });
  await waitFor(() => expect(value.view.queryByTestId('pptx-proposals-panel')).toBeNull());
  expect(button.getAttribute('aria-expanded')).toBe('false');
});

test('a read-only viewer has no proposals button', async () => {
  const value = await open({ readOnly: true });
  expect(value.view.getByTestId('pptx-present')).toBeDefined();
  expect(value.view.queryByTestId('pptx-proposals-button')).toBeNull();
});

test('the proposals button stays hidden when the worker has no proposals support', async () => {
  const propose = Object.getOwnPropertyDescriptor(PptxDocument.prototype, 'proposeJson')!;
  const strip = spyOn(PptxDocument.prototype, 'proposalSlideStripJson');
  const preview = spyOn(PptxRenderer.prototype, 'layoutProposalDiffSlideByIdJson');
  Object.defineProperty(PptxDocument.prototype, 'proposeJson', { configurable: true, value: undefined });
  try {
    const value = await open();
    expect(value.view.getByTestId('pptx-present')).toBeDefined();
    expect(value.view.queryByTestId('pptx-proposals-button')).toBeNull();
    const access = await value.api.handleAsync();
    expect(await access.proposalSlideStrip('unsupported')).toEqual([]);
    await expect(access.layoutProposalDiffSlideById('unsupported', value.client.state.slides[0]!.id)).rejects.toThrow('Proposal slide previews are unavailable');
    expect(strip).not.toHaveBeenCalled();
    expect(preview).not.toHaveBeenCalled();
  } finally {
    Object.defineProperty(PptxDocument.prototype, 'proposeJson', propose);
    strip.mockRestore();
    preview.mockRestore();
  }
});

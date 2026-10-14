import { open, replyControl, story, type } from './peerlessTestHarness';
import { expect, mock, spyOn, test } from 'bun:test';
import type { PresentationHandle } from '@betteroffice/pptx';
const { act, fireEvent, waitFor } = await import('@testing-library/react');
type ReviewAccess = { [K in 'propose' | 'previewProposal' | 'layoutProposalDiffSlide' | 'acceptProposal' | 'rejectProposal']:
  (...args: Parameters<PresentationHandle[K]>) => Promise<ReturnType<PresentationHandle[K]>> };

function composition(input: HTMLElement, kind: string, text = ''): void {
  const event = new Event(kind, { bubbles: true });
  Object.defineProperty(event, 'data', { value: text });
  input.dispatchEvent(event);
}

test('proposal diff commits composition before removing the input', async () => {
  const value = await open();
  const target = await story(value.api);
  const access = await value.api.handleAsync();
  let id = '';
  await act(async () => {
    id = (await access.propose('reviewer', 'Update', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Updated' }])).id;
    await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 });
    await value.api.commands.execute('proposalSelect', { proposalId: id });
  });
  await act(async () => {
    const input = value.view.getByLabelText('Slide text');
    composition(input, 'compositionstart');
    composition(input, 'compositionupdate', '漢');
    await value.api.flushPendingInput();
  });
  await act(async () => { expect(await value.api.commands.execute('proposalDiff', { enabled: true })).toMatchObject({ ok: true }); });
  expect(value.view.queryByLabelText('Slide text')).toBeNull();
  await act(async () => { await value.api.commands.execute('proposalDiff', { enabled: false }); });
  await act(async () => { type(value.view.getByLabelText('Slide text'), 'NEXT'); await value.api.flushPendingInput(); });
  expect((await story(value.api)).text).toBe(`漢NEXT${target.text}`);
});

test('removing the selected proposal through the API restores editing', async () => {
  const value = await open();
  const target = await story(value.api);
  const access = await value.api.handleAsync();
  let id = '';
  await act(async () => {
    id = (await access.propose('reviewer', 'Update', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Updated' }])).id;
    await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 });
    await value.api.commands.execute('proposalSelect', { proposalId: id });
  });
  await act(async () => { await value.api.commands.execute('proposalDiff', { enabled: true }); });
  expect(value.view.queryByLabelText('Slide text')).toBeNull();
  await act(async () => { await access.rejectProposal(id); });
  await waitFor(() => expect(value.view.queryByLabelText('Slide text')).not.toBeNull());
  expect(value.view.queryByTestId('pptx-proposal-diff')).toBeNull();
  expect(value.api.commands.getState('proposalSelect').value).toBeNull();
  expect(document.activeElement).toBe(value.view.getByLabelText('Slide text'));
  await act(async () => { type(value.view.getByLabelText('Slide text'), 'NEXT'); await value.api.flushPendingInput(); });
  expect((await story(value.api)).text).toBe(`NEXT${target.text}`);
});

test('proposal preview navigates to its slide before selecting the diff', async () => {
  const error = mock(() => {});
  const value = await open({ onError: error });
  const access = await value.api.handleAsync();
  const content = await access.readContent();
  if (!content.ok) throw new Error(content.failure.message);
  const target = content.stories.find((story) => story.slideId === value.client.state.slides[1]!.id)!;
  await act(async () => {
    await access.propose('reviewer', 'Second slide', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Updated' }]);
    await value.api.commands.execute('proposalsPanel', { open: true });
  });
  act(() => { fireEvent.click(value.view.getByRole('button', { name: 'Preview changes' })); });
  await waitFor(() => expect(value.client.state.activeSlideId).toBe(target.slideId));
  await waitFor(() => expect(value.view.getByTestId('pptx-proposal-diff')).toBeDefined());
  expect(error).not.toHaveBeenCalled();
});

test('proposal preview reports a selection refusal', async () => {
  const error = mock(() => {});
  const value = await open({ onError: error });
  const access = await value.api.handleAsync();
  const target = await story(value.api);
  await act(async () => {
    await access.propose('reviewer', 'Update', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Updated' }]);
    await value.api.commands.execute('proposalsPanel', { open: true });
  });
  const execute = spyOn(value.api.commands, 'execute').mockResolvedValue({ ok: false,
    failure: { code: 'target-changed', message: 'Proposal target changed' } });
  try {
    act(() => { fireEvent.click(value.view.getByRole('button', { name: 'Preview changes' })); });
    await waitFor(() => expect(error).toHaveBeenCalledWith(expect.objectContaining({ message: 'Proposal target changed' })));
    expect(value.view.getByRole('alert').textContent).toBe('Proposal target changed');
    expect(execute).toHaveBeenCalledTimes(1);
  } finally { execute.mockRestore(); }
});

test('PNG export resolves the active slide after navigation acknowledgement', async () => {
  const value = await open();
  const slideId = value.client.state.slides[1]!.id;
  const thumbnail = spyOn(value.client, 'thumbnail');
  const encode = spyOn(HTMLCanvasElement.prototype, 'toBlob').mockImplementation((callback) => callback(new Blob([], { type: 'image/png' })));
  const createUrl = spyOn(URL, 'createObjectURL').mockReturnValue('blob:slide');
  const revokeUrl = spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
  const download = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
  replyControl.hold();
  const navigation = value.api.goToSlide(2);
  try {
    await waitFor(() => expect(replyControl.count()).toBe(1));
    const exporting = value.api.commands.execute('exportPng', null);
    replyControl.release();
    await act(async () => { await navigation; expect(await exporting).toMatchObject({ ok: true }); });
    expect(thumbnail.mock.calls[thumbnail.mock.calls.length - 1]?.[0]).toBe(slideId);
  } finally {
    replyControl.release(); thumbnail.mockRestore(); encode.mockRestore(); createUrl.mockRestore(); revokeUrl.mockRestore(); download.mockRestore();
  }
});

test('async adjustment proposals are admitted and accepted', async () => {
  const value = await open();
  await act(async () => { await value.api.commands.execute('tool', { value: 'shape:roundRect' }); });
  await act(async () => {
    fireEvent.pointerDown(value.canvas, value.point(320, 200));
    fireEvent.pointerUp(value.stage, value.point(320, 200));
    await value.api.flushPendingInput();
  });
  const selected = value.client.state.selection!;
  const access = await value.api.handleAsync();
  await act(async () => {
    const proposal = await access.propose('reviewer', 'Adjust corners', [{ type: 'setShapeAdjust', slideId: selected.slideId,
      shapeId: selected.shapeId, adjustments: { adj: 0.32 } }]);
    expect(proposal.changes).toHaveLength(1);
    expect((await access.acceptProposal(proposal.id)).applied).toBe(true);
  });
  expect(value.client.state.editState.object?.adjustments.adj).toBe(0.32);
});

test('async_review_export', async () => {
  const value = await open();
  const target = await story(value.api);
  const access = await value.api.handleAsync();
  const review = access as unknown as ReviewAccess;
  let id = '';
  await act(async () => { id = (await review.propose('reviewer', 'Update', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Updated' }])).id; });
  expect((await review.previewProposal(id)).proposal.id).toBe(id);
  expect((await review.layoutProposalDiffSlide(id, 0)).frame.primitives.length).toBeGreaterThan(0);
  await act(async () => { expect(await value.api.commands!.execute('proposalsPanel', { open: true })).toMatchObject({ ok: true }); });
  expect(value.view.getByTestId('pptx-proposals-panel')).toBeDefined();
  await act(async () => { expect(await value.api.commands!.execute('proposalAccept', { proposalId: id })).toMatchObject({ ok: true }); });
  expect((await story(value.api)).text.startsWith('Updated')).toBe(true);
  const markdown = await access.exportMarkdown();
  expect(markdown.ok).toBe(true);
  if (!markdown.ok) throw new Error(markdown.failure.message);
  expect(markdown.content.markdown).toContain('Updated');
  const structured = await access.exportStructured();
  expect(structured.ok).toBe(true);
  if (!structured.ok) throw new Error(structured.failure.message);
  expect(structured.content.slides).toHaveLength(2);
  const png = new Blob([new Uint8Array([137, 80, 78, 71])], { type: 'image/png' });
  const encode = spyOn(HTMLCanvasElement.prototype, 'toBlob').mockImplementation((callback) => callback(png));
  const createUrl = spyOn(URL, 'createObjectURL').mockReturnValue('blob:slide');
  const revokeUrl = spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
  const download = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
  try {
    await act(async () => { expect(await value.api.commands!.execute('exportPng', null)).toMatchObject({ ok: true }); });
    expect(encode).toHaveBeenCalledWith(expect.any(Function), 'image/png');
    expect(createUrl).toHaveBeenCalledWith(png);
    expect(download).toHaveBeenCalledTimes(1);
    expect(revokeUrl).toHaveBeenCalledWith('blob:slide');
  } finally { encode.mockRestore(); createUrl.mockRestore(); revokeUrl.mockRestore(); download.mockRestore(); }
  await act(async () => { expect(await value.api.commands!.execute('slideshow', null)).toMatchObject({ ok: true }); });
  expect(value.view.getByRole('dialog')).toBeDefined();
  await act(async () => { value.view.rerender(value.ui(true)); await value.api.flushPendingInput(); });
  await expect(review.acceptProposal(id)).rejects.toThrow();
});

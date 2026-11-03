import { deck, open, replyControl, story, type } from './peerlessTestHarness';
import { expect, mock, spyOn, test } from 'bun:test';
import type { PresentationHandle } from '@betteroffice/pptx';
import JSZip from 'jszip';
import * as frameSources from './peerlessFrameSource';
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

test('stale notes preview shows the current notes before enabling force acceptance', async () => {
  const value = await open();
  const access = await value.api.handleAsync();
  const slideId = value.client.state.slides[0]!.id;
  await act(async () => {
    await access.setSlideNotes(slideId, 'Original notes');
    await access.propose('reviewer', 'Update notes', [{ type: 'setSlideNotes', slideId, text: 'Proposed notes' }]);
    await access.setSlideNotes(slideId, 'Newer notes');
    await value.api.commands.execute('proposalsPanel', { open: true });
  });
  expect(value.view.queryByTestId('pptx-proposal-force')).toBeNull();
  act(() => { fireEvent.click(value.view.getByRole('button', { name: 'Preview changes' })); });
  await waitFor(() => expect(value.view.getByTestId('pptx-proposal-notes-diff').querySelector('del')?.textContent).toBe('Newer notes'));
  expect(value.view.getByTestId('pptx-proposal-notes-diff').querySelector('ins')?.textContent).toBe('Proposed notes');
  expect(value.view.getByTestId('pptx-proposal-force').textContent).toBe('Apply updated proposal');
  expect(value.view.getByTestId('pptx-proposal').textContent).toContain('Applying will override the changed targets');
  await act(async () => { fireEvent.click(value.view.getByTestId('pptx-proposal-force')); });
  await waitFor(() => expect(value.view.queryByTestId('pptx-proposal') === null).toBe(true));
  expect((await access.snapshot()).slides[0]!.notes).toBe('Proposed notes');
});

test('force acceptance refreshes notes changed since preview without applying them', async () => {
  const value = await open();
  const access = await value.api.handleAsync();
  const slideId = value.client.state.slides[0]!.id;
  await act(async () => {
    await access.propose('reviewer', 'Update notes', [{ type: 'setSlideNotes', slideId, text: 'Proposed notes' }]);
    await access.setSlideNotes(slideId, 'Newer notes');
    await value.api.commands.execute('proposalsPanel', { open: true });
  });
  act(() => { fireEvent.click(value.view.getByRole('button', { name: 'Preview changes' })); });
  await waitFor(() => expect(value.view.getByTestId('pptx-proposal-force')).toBeDefined());
  const read = value.client.access!;
  let changed = false;
  const refresh = spyOn(value.client, 'access').mockImplementation((async (method, ...args) => {
    if (method === 'previewProposal' && !changed) {
      changed = true;
      await access.setSlideNotes(slideId, 'Changed again');
    }
    return (read as (...args: unknown[]) => Promise<unknown>)(method, ...args);
  }) as typeof read);
  try {
    await act(async () => { fireEvent.click(value.view.getByTestId('pptx-proposal-force')); });
    await waitFor(() => expect(value.view.getAllByRole('alert').some((alert) => alert.textContent?.includes('changed again'))).toBe(true));
  } finally { refresh.mockRestore(); }
  await waitFor(() => expect(value.view.getByTestId('pptx-proposal-notes-diff').querySelector('del')?.textContent).toBe('Changed again'));
  expect((await access.snapshot()).slides[0]!.notes).toBe('Changed again');
  expect(await access.listProposals()).toHaveLength(1);
  expect(value.view.getAllByRole('alert').some((alert) => alert.textContent?.includes('changed again'))).toBe(true);
  await waitFor(() => expect(value.view.queryByTestId('pptx-proposal-force') !== null).toBe(true));
  await act(async () => { fireEvent.click(value.view.getByTestId('pptx-proposal-force')); });
  await waitFor(() => expect(value.view.queryByTestId('pptx-proposal') === null).toBe(true));
  expect((await access.snapshot()).slides[0]!.notes).toBe('Proposed notes');
});

test('structural proposals show ghost strip entries and accept from the canvas review', async () => {
  const value = await open();
  const access = await value.api.handleAsync();
  let id = '';
  await act(async () => {
    id = (await access.propose('reviewer', 'New slide', [
      { type: 'insertSlide', ref: 'draft', afterSlideId: null },
      { type: 'setSlideNotes', slideId: '@draft', text: 'Draft' },
      { type: 'deleteSlide', slideId: value.client.state.slides[1]!.id },
    ])).id;
    await value.api.commands.execute('proposalsPanel', { open: true });
    await value.api.commands.execute('proposalSelect', { proposalId: id });
    await value.api.commands.execute('proposalDiff', { enabled: true });
  });
  const ghost = await waitFor(() => value.view.getByTestId('pptx-proposal-slide-inserted'));
  expect(ghost.style.borderStyle).toBe('dashed');
  expect(value.view.getByTestId('pptx-proposal-slide-deleted').style.opacity).toBe('0.4');
  await act(async () => { fireEvent.click(ghost); });
  await waitFor(() => expect(value.view.getByTestId('pptx-proposal-diff').getAttribute('data-slide-id')).toBe(ghost.getAttribute('data-slide-id')));
  await act(async () => { fireEvent.click(value.view.getByTestId('pptx-proposal-accept')); });
  await waitFor(async () => expect(await access.listProposals()).toHaveLength(0));
  expect((await access.snapshot()).slides[0]!.notes).toBe('Draft');
});

test('inserted layout background images paint in ghost thumbnails and canvas previews', async () => {
  const zip = await JSZip.loadAsync(await deck());
  const imageId = 'ppt/media/layout-background.png';
  const layout = 'ppt/slideLayouts/slideLayout1.xml';
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
  zip.file(layout, `<p:sldLayout xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}" type="blank"><p:cSld name="Picture background"><p:bg><p:bgPr><a:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg><p:spTree/></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<Relationships xmlns="${rels}"><Relationship Id="picture" Type="${r}/image" Target="../media/layout-background.png"/></Relationships>`);
  zip.file(imageId, Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg=='), char => char.charCodeAt(0)));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', `<Default Extension="png" ContentType="image/png"/><Override PartName="/${layout}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/></Types>`));
  const bitmap = { width: 1, height: 1, close: mock(() => {}) } as unknown as ImageBitmap;
  const originalBitmap = Object.getOwnPropertyDescriptor(globalThis, 'createImageBitmap');
  Object.defineProperty(globalThis, 'createImageBitmap', { configurable: true, value: mock(async () => bitmap) });
  const draws = new Map<HTMLCanvasElement, unknown[]>();
  const context = HTMLCanvasElement.prototype.getContext;
  const paint = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation((function (this: HTMLCanvasElement, ...args: unknown[]) {
    const result = Reflect.apply(context, this, args) as CanvasRenderingContext2D | null;
    if (!result) return result;
    const canvas = this;
    return new Proxy(result, { get(target, key) {
      if (key === 'drawImage') return (image: unknown) => { draws.set(canvas, [...(draws.get(canvas) ?? []), image]); };
      return Reflect.get(target, key);
    } });
  }) as HTMLCanvasElement['getContext']);
  const resolve = mock((id: string) => {});
  const releases: ReturnType<typeof mock>[] = [];
  const load = frameSources.peerlessProposalFrameSource;
  const source = spyOn(frameSources, 'peerlessProposalFrameSource').mockImplementation(async (...args) => {
    const result = await load(...args);
    const release = mock(result.release);
    releases.push(release);
    return { ...result, resolveImage: (id: string) => { resolve(id); return result.resolveImage(id); }, release };
  });
  try {
    const value = await open({ file: await zip.generateAsync({ type: 'uint8array' }) });
    const access = await value.api.handleAsync();
    await act(async () => {
      const proposal = await access.propose('reviewer', 'Picture layout', [{ type: 'insertSlide', afterSlideId: null, layoutPartPath: layout }]);
      await value.api.commands.execute('proposalSelect', { proposalId: proposal.id });
      await value.api.commands.execute('proposalDiff', { enabled: true });
    });
    const ghost = await waitFor(() => value.view.getByTestId('pptx-proposal-slide-inserted'));
    const paintedImage = (canvas: HTMLCanvasElement) => (draws.get(canvas) ?? []).some(buffer =>
      buffer instanceof HTMLCanvasElement && draws.get(buffer)?.includes(bitmap));
    await waitFor(() => expect(paintedImage(ghost.querySelector('canvas')!)).toBe(true));
    expect(resolve).toHaveBeenCalledWith(imageId);
    resolve.mockClear();
    await act(async () => { fireEvent.click(ghost); });
    await waitFor(() => expect(paintedImage(value.view.getByTestId('pptx-proposal-diff') as HTMLCanvasElement)).toBe(true));
    expect(resolve).toHaveBeenCalledWith(imageId);
    await waitFor(() => { for (const release of releases) expect(release).toHaveBeenCalledTimes(1); });
  } finally {
    source.mockRestore(); paint.mockRestore();
    if (originalBitmap) Object.defineProperty(globalThis, 'createImageBitmap', originalBitmap);
    else Reflect.deleteProperty(globalThis, 'createImageBitmap');
  }
});

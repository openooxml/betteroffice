import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { initWasm, openPresentation } from '@betteroffice/pptx';
import type { ProposalDiffSlide, TextParagraphDraft } from '@betteroffice/pptx';
import { ProposalCanvasOverlay, ProposalShapeOutlines, useProposalCanvas } from './ProposalCanvas';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, fireEvent, render, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

let fixture: Uint8Array;
let font: Uint8Array;
beforeAll(async () => {
  const [wasm, deck, face] = await Promise.all([
    readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)),
    readFile(new URL('../../../../crates/pptx-edit/tests/fixtures/proposal-shapes.pptx', import.meta.url)),
    readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)),
  ]);
  await initWasm(wasm);
  fixture = deck;
  font = face;
});

test.each([
  { paragraphs: [{ bullet: true, runs: [{ text: 'First', style: {} }] }, { runs: [{ text: 'Second', style: {} }] }] },
  { paragraphs: [{ bullet: true, runs: [{ text: 'alpha', style: {} }] }] },
  { paragraphs: [{ alignment: 'ctr', runs: [{ text: 'alpha', style: {} }] }] },
] satisfies { paragraphs: TextParagraphDraft[] }[])('canvas review paints paragraph drafts %# and exposes their changed shape', async ({ paragraphs }) => {
  const handle = openPresentation(fixture, { fonts: [{ family: 'Liberation Sans', bytes: font }] });
  const paintedText: string[] = [];
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return new Proxy({ canvas: this, measureText: (text: string) => ({ width: text.length * 8 }),
      fillText: (text: string) => paintedText.push(text) }, {
      get: (target, key) => Reflect.get(target, key) ?? (() => {}),
    }) as unknown as CanvasRenderingContext2D;
  } as unknown as typeof HTMLCanvasElement.prototype.getContext);
  try {
    const before = handle.snapshot();
    const shape = before.slides[0]!.shapes[0]!;
    handle.propose('agent', null, [{ type: 'setStoryParagraphs', storyId: shape.textStories[0]!.id, paragraphs }]);
    const review = renderHook(() => useProposalCanvas(handle, handle.listProposals(), before, 0));
    expect(review.result.current.error).toBeNull();
    const diff = review.result.current.diff!;
    expect(diff.shapeChanges).toContainEqual({ shapeId: shape.id, kind: 'changed' });
    expect(diff.textChanges.length > 0).toBe(paragraphs.length > 1);
    let painted: ProposalDiffSlide | null = null;
    let target: string | null = null;
    const view = render(<>
      <ProposalCanvasOverlay diff={diff} current={before} frame={diff.frame} slideIndex={0} scale={1}
        resolveImage={() => null} onTarget={() => {}} onPainted={value => { painted = value; }} />
      <ProposalShapeOutlines diff={diff} current={before} frame={diff.frame} scale={1} onTarget={(_slide, shapeId) => { target = shapeId; }} />
    </>);
    await waitFor(() => expect(painted).toBe(diff));
    expect(view.getByRole('img').getAttribute('data-ready')).toBe('true');
    const outline = view.getByTestId('pptx-proposal-canvas-target');
    expect(outline.getAttribute('data-shape-change')).toBe('changed');
    fireEvent.click(outline);
    expect(target).toBe(shape.id);
    if (paragraphs.length > 1) {
      expect(paintedText.join('')).toContain('First');
      expect(paintedText.join('')).toContain('Second');
      expect(diff.textChanges.some(change => change.kind === 'deletion')).toBe(true);
    }
    expect(handle.snapshot()).toEqual(before);
  } finally {
    cleanup();
    context.mockRestore();
    handle.dispose();
  }
});

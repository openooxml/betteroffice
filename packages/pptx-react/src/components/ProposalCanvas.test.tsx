import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { initWasm, openPresentation } from '@betteroffice/pptx';
import type { ProposalDiffSlide, TextParagraphDraft } from '@betteroffice/pptx';
import { createT, en, de, deepMerge, type LocaleStrings } from '@betteroffice/pptx-i18n';
import { ProposalChangeDetails } from './ProposalDetails';
import { ProposalSlideStrip } from './ProposalSlideStrip';
import { ProposalCanvasOverlay, useProposalCanvas } from './ProposalCanvas';
import { LocaleProvider } from '../i18n/LocaleContext';
import { useProposalTranslation } from '../i18n/useProposalTranslation';

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
    const onTarget = mock((_slideId: string, _shapeId: string) => {});
    const view = render(
      <ProposalCanvasOverlay diff={diff} current={before} frame={diff.frame} slideIndex={0} scale={1}
        resolveImage={() => null} onTarget={onTarget} onPainted={value => { painted = value; }} />
    );
    await waitFor(() => expect(painted).toBe(diff));
    expect(view.getByRole('img').getAttribute('data-ready')).toBe('true');
    const outline = view.getByTestId('pptx-proposal-canvas-target');
    expect(outline.getAttribute('data-shape-change')).toBe('changed');
    fireEvent.click(outline);
    expect(onTarget).toHaveBeenCalledWith(before.slides[0]!.id, shape.id);
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

test.each(['setSlideLayout', 'setSlideBackground'] as const)('canvas and strip preview %s and review its translated label', async type => {
  const bytes = await readFile(new URL('../../../../crates/pptx-edit/tests/fixtures/proposal-slide-props.pptx', import.meta.url));
  const handle = openPresentation(bytes, { fonts: [{ family: 'Liberation Sans', bytes: font }] });
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return new Proxy({ canvas: this, measureText: (text: string) => ({ width: text.length * 8 }) }, {
      get: (target, key) => Reflect.get(target, key) ?? (() => {}),
    }) as unknown as CanvasRenderingContext2D;
  } as unknown as typeof HTMLCanvasElement.prototype.getContext);
  try {
    const before = handle.snapshot();
    const slide = before.slides[0]!;
    const proposal = handle.propose('agent', null, [type === 'setSlideLayout'
      ? { type, slideId: slide.id, layoutPartPath: 'ppt/slideLayouts/slideLayout2.xml' }
      : { type, slideId: slide.id, background: { type: 'solid', color: { rgb: '123456' } } }]);
    const review = renderHook(() => useProposalCanvas(handle, handle.listProposals(), before, 0));
    expect(review.result.current.error).toBeNull();
    const diff = review.result.current.diff!;
    expect(JSON.stringify(diff.frame)).toContain('#123456');
    expect(diff.proposal.changes[0]!.shapeId).toBeNull();
    const entries = handle.proposalSlideStrip(proposal.id);
    let selected = '';
    let painted: ProposalDiffSlide | null = null;
    let decoded = 0;
    const decodeImage = async () => { decoded++; return document.createElement('canvas'); };
    const view = render(<>
      <ProposalSlideStrip entries={entries} selectedSlideId={slide.id} load={() => diff}
        decodeImage={decodeImage} onSelect={id => { selected = id; }} />
      <ProposalCanvasOverlay diff={diff} current={before} frame={diff.frame} slideIndex={0} scale={1}
        resolveImage={() => null} decodeImage={decodeImage} onTarget={() => {}} onPainted={value => { painted = value; }} />
      <ProposalChangeDetails change={diff.proposal.changes[0]!} proposal={diff.proposal} t={createT(en)} />
    </>);
    await waitFor(() => expect(painted).toBe(diff));
    expect(view.getByText(type === 'setSlideLayout' ? 'Layout changed' : 'Background changed')).toBeTruthy();
    fireEvent.click(view.getByTestId('pptx-proposal-slide-changed'));
    expect(selected).toBe(slide.id);
    if (type === 'setSlideLayout') {
      expect(diff.media?.some(asset => asset.assetId === 'ppt/media/proposal-pixel.png')).toBe(true);
      expect(decoded).toBeGreaterThan(0);
    }
    expect(handle.snapshot()).toEqual(before);
  } finally { cleanup(); context.mockRestore(); handle.dispose(); }
});

test('slide review labels resolve created slide refs and keep both property changes', () => {
  const change: import('@betteroffice/pptx').ProposalChange = {
    slideId: 'created-slide', shapeId: null, kind: 'content', before: null, after: null, oldText: '', newText: '',
  };
  const proposal: import('@betteroffice/pptx').Proposal = {
    id: 'p1', agentId: 'agent', note: null, staleTargets: [], refs: { new: change.slideId }, changes: [change], edits: [
      { type: 'setSlideLayout', slideId: '@new', layoutPartPath: 'layout.xml' },
      { type: 'setSlideBackground', slideId: '@new', background: null },
    ],
  };
  const view = render(<ProposalChangeDetails change={change} proposal={proposal} t={createT(deepMerge(en, de) as LocaleStrings, 'de')} textOnly />);
  expect(view.getByText('Layout geändert')).toBeTruthy();
  expect(view.getByText('Hintergrund geändert')).toBeTruthy();
});

test.each([
  ['insertTableRow', 'Row inserted'], ['deleteTableRow', 'Row deleted'],
  ['insertTableColumn', 'Column inserted'], ['deleteTableColumn', 'Column deleted'],
] as const)('canvas previews %s table structure and loads its review label lazily', async (type, label) => {
  const zip = await JSZip.loadAsync(await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/table-basic.pptx', import.meta.url)));
  const path = 'ppt/slides/slide1.xml';
  zip.file(path, (await zip.file(path)!.async('string')).replace(/ (gridSpan|rowSpan|hMerge|vMerge)="[^"]*"/gu, ''));
  const handle = openPresentation(await zip.generateAsync({ type: 'uint8array' }), { fonts: [{ family: 'Liberation Sans', bytes: font }] });
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return new Proxy({ canvas: this, measureText: (text: string) => ({ width: text.length * 8 }) }, {
      get: (target, key) => Reflect.get(target, key) ?? (() => {}),
    }) as unknown as CanvasRenderingContext2D;
  } as unknown as typeof HTMLCanvasElement.prototype.getContext);
  try {
    const before = handle.snapshot();
    const slide = before.slides[0]!;
    const shape = slide.shapes[0]!;
    handle.propose('agent', null, [type === 'insertTableRow' || type === 'insertTableColumn'
      ? { type, slideId: slide.id, shapeId: shape.id, index: 0, position: 'after', cells: ['Review cell'] }
      : { type, slideId: slide.id, shapeId: shape.id, index: 0 }]);
    const review = renderHook(() => useProposalCanvas(handle, handle.listProposals(), before, 0));
    expect(review.result.current.error).toBeNull();
    const diff = review.result.current.diff!;
    expect(diff.shapeChanges).toEqual([{ shapeId: shape.id, kind: 'changed' }]);
    const table = diff.snapshot.slides[0]!.shapes[0]!.graphic as { rows: unknown[]; grid: unknown[] };
    expect(table.rows.length).toBe(type === 'insertTableRow' ? 4 : type === 'deleteTableRow' ? 2 : 3);
    expect(table.grid.length).toBe(type === 'insertTableColumn' ? 4 : type === 'deleteTableColumn' ? 2 : 3);
    function Details() {
      const { t } = useProposalTranslation();
      return <ProposalChangeDetails change={diff.proposal.changes[0]!} proposal={diff.proposal} t={t} />;
    }
    let painted: ProposalDiffSlide | null = null;
    const view = render(<LocaleProvider>
      <ProposalCanvasOverlay diff={diff} current={before} frame={diff.frame} slideIndex={0} scale={1}
        resolveImage={() => null} onTarget={() => {}} onPainted={value => { painted = value; }} />
      <Details />
    </LocaleProvider>);
    await waitFor(() => expect(painted).toBe(diff));
    expect(view.getByText(label)).toBeTruthy();
    expect(view.getByTestId('pptx-proposal-canvas-target').getAttribute('data-shape-change')).toBe('changed');
    expect(handle.snapshot()).toEqual(before);
  } finally { cleanup(); context.mockRestore(); handle.dispose(); }
});

test('table review labels resolve slide and shape refs, preserve each edit and use locale translations', () => {
  const change: import('@betteroffice/pptx').ProposalChange = {
    slideId: 'created-slide', shapeId: 'created-table', kind: 'content', before: null, after: null, oldText: '', newText: '',
  };
  const proposal: import('@betteroffice/pptx').Proposal = {
    id: 'p1', agentId: 'agent', note: null, staleTargets: [], refs: { slide: change.slideId, table: change.shapeId! }, changes: [change], edits: [
      { type: 'insertTableRow', slideId: '@slide', shapeId: '@table', index: 0, position: 'before' },
      { type: 'insertTableRow', slideId: '@slide', shapeId: '@table', index: 1, position: 'after' },
      { type: 'deleteTableRow', slideId: '@slide', shapeId: '@table', index: 0 },
      { type: 'insertTableColumn', slideId: '@slide', shapeId: '@table', index: 0, position: 'after' },
      { type: 'deleteTableColumn', slideId: '@slide', shapeId: '@table', index: 0 },
      { type: 'insertTableRow', slideId: 'other-slide', shapeId: '@table', index: 0, position: 'before' },
      { type: 'deleteTableRow', slideId: '@slide', shapeId: 'other-table', index: 0 },
    ],
  };
  function Details() {
    const { t } = useProposalTranslation();
    return <ProposalChangeDetails change={change} proposal={proposal} t={t} textOnly />;
  }
  const view = render(<LocaleProvider i18n={de}><Details /></LocaleProvider>);
  expect(view.getAllByText('Zeile eingefügt')).toHaveLength(2);
  expect(view.getAllByText('Zeile gelöscht')).toHaveLength(1);
  expect(view.getByText('Spalte eingefügt')).toBeTruthy();
  expect(view.getByText('Spalte gelöscht')).toBeTruthy();
});

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, spyOn, test } from 'bun:test';
import * as pptx from '@betteroffice/pptx';
import type { ProposalDiffSlide, ProposalStripEntry, SlideDisplayList } from '@betteroffice/pptx';
import { UNAVAILABLE_PPTX_COMMANDS } from '../commands/createPptxCommandStore';
import { ProposalSlideStrip } from '../components/ProposalSlideStrip';
import { PeerlessPresentationOverlay } from './PeerlessPresentationOverlay';
import { PeerlessReview } from './PeerlessReview';
import { PeerlessSlideStrip } from './PeerlessSlideStrip';
import type { PeerlessPresentation } from './usePeerlessPresentation';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render, waitFor } = await import('@testing-library/react');
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const frame: SlideDisplayList = { contractVersion: 1, width: 320, height: 180, primitives: [] };
const diff: ProposalDiffSlide = {
  frame, textChanges: [], snapshot: { widthEmu: 0, heightEmu: 0, slides: [] },
  proposal: { id: 'proposal', agentId: 'reviewer', note: null, edits: [], changes: [], staleTargets: [] },
};
const entries: ProposalStripEntry[] = [{ slideId: 'slide', status: 'changed', liveIndex: 0, previewIndex: 0 }];
const fontCoverage = {
  prepare: async (family: string) => family,
  thumbnails: { prepare: async (family: string) => family },
};
const resolveImage = () => null;
const idle = () => () => {};
const commands = UNAVAILABLE_PPTX_COMMANDS;
const run = {
  fontCoverage, current: true, require() {},
  getVisualSnapshot: () => null, subscribeVisual: idle,
  client: {
    state: { slides: [{ id: 'slide' }], activeSlideId: 'slide' },
    access: async (method: string) => method === 'proposalSlideStrip' ? entries : diff,
    thumbnail: async () => ({
      status: 'frame', frame: { displayList: new TextEncoder().encode(JSON.stringify(frame)), media: [] },
    }),
  },
  store: {
    fence: async () => {}, getSnapshot: () => ({ editState: { readOnly: true }, thumbnails: new Map() }),
    subscribeStructure: idle, subscribePresentation: idle, subscribeEditState: idle,
  },
  images: { pin: () => ({ ready: Promise.resolve(), resolve: resolveImage, release() {} }) },
} as unknown as PeerlessPresentation;

let canvas: ReturnType<typeof spyOn<HTMLCanvasElement, 'getContext'>>;
let paint: ReturnType<typeof spyOn<typeof pptx, 'paintSlide'>>;
beforeEach(() => {
  canvas = spyOn(HTMLCanvasElement.prototype, 'getContext')
    .mockImplementation((() => ({ drawImage() {} })) as unknown as typeof HTMLCanvasElement.prototype.getContext);
  paint = spyOn(pptx, 'paintSlide').mockResolvedValue(undefined);
});
afterEach(() => {
  cleanup();
  paint.mockRestore();
  canvas.mockRestore();
});

test('the presentation overlay passes font coverage to its painter', async () => {
  render(<PeerlessPresentationOverlay fontCoverage={fontCoverage}
    frameSource={async () => ({ frame, resolveImage, release() {} })}
    slideCount={1} startIndex={0} label="Presentation" counterLabel={() => '1 / 1'}
    exitLabel="Exit" previousLabel="Previous" nextLabel="Next" onExit={() => {}} onError={() => {}} />);
  await waitFor(() => expect(paint.mock.calls[0]?.[4]?.fontCoverage).toBe(fontCoverage));
});

test.each(['proposal diff', 'presentation'] as const)('peerless review passes font coverage to the %s painter', async (path) => {
  const presenting = path === 'presentation';
  render(<PeerlessReview run={run} commands={commands} proposals={[diff.proposal]} onError={() => {}} exit={() => {}}
    diffHost={{ current: document.createElement('div') }}
    state={{ open: false, selectedId: diff.proposal.id, diff: !presenting, presenting }} />);
  await waitFor(() => expect(paint.mock.calls[0]?.[4]?.fontCoverage).toBe(fontCoverage));
});

test('proposal thumbnails from the peerless strip paint with thumbnail font coverage', async () => {
  render(<PeerlessSlideStrip run={run} commands={commands} previewFit={() => null} current={() => true}
    proposalId={diff.proposal.id} onPreviewSlide={() => {}} onNavigate={() => {}} />);
  await waitFor(() => expect(paint.mock.calls[0]?.[4]?.fontCoverage).toBe(fontCoverage.thumbnails));
});

test('the proposal slide strip passes font coverage to its thumbnail painter', async () => {
  render(<ProposalSlideStrip entries={entries} load={() => diff} onSelect={() => {}} fontCoverage={fontCoverage} />);
  await waitFor(() => expect(paint.mock.calls[0]?.[4]?.fontCoverage).toBe(fontCoverage));
});

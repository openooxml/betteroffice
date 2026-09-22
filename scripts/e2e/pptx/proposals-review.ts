import { expect } from 'bun:test';

import { isProposalsAvailable } from '../../../packages/pptx/src/wasm/loader';
import { firstParagraphText, firstStory, profiledLayout } from './context';
import type { PptxScenario } from './context';

const SLIDE = 0;

export const proposalsReview: PptxScenario = {
  name: 'proposals-review',
  description:
    'An agent proposes a text replacement and a fill change, the reviewer previews the proposal and renders its diff slide without touching the deck, then accepts one proposal and rejects the other.',
  participants: ['web', 'agent'],
  requires: () =>
    isProposalsAvailable()
      ? undefined
      : 'the core was built without the proposals api',
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const slide = recorder.op('snapshot', () => handle.snapshot()).slides[
      SLIDE
    ];
    const story = firstStory(slide.shapes);
    const original = firstParagraphText(story);
    const target = slide.shapes.find((shape) =>
      shape.textStories.some((entry) => entry.id === story.id)
    )!;

    const rewrite = recorder.op(
      'propose:rewrite',
      () =>
        handle.propose('agent-a', 'shorter headline', [
          {
            type: 'replaceText',
            storyId: story.id,
            start: 0,
            end: original.length,
            text: 'Proposed headline',
          },
        ]),
      undefined,
      { actor: 'agent' }
    );
    const recolour = recorder.op(
      'propose:recolour',
      () =>
        handle.propose('agent-b', 'brand fill', [
          {
            type: 'setShapeFill',
            slideId: slide.id,
            shapeId: target.id,
            color: '#8811aa',
          },
        ]),
      undefined,
      { actor: 'agent' }
    );
    const change = rewrite.changes.find((entry) => entry.shapeId === target.id);
    expect(change?.slideId).toBe(slide.id);
    expect(change?.oldText).toContain(original);
    expect(change?.newText).toContain('Proposed headline');
    expect(
      recorder
        .op('listProposals', () => handle.listProposals())
        .map((proposal) => proposal.id)
        .sort()
    ).toEqual([rewrite.id, recolour.id].sort());
    expect(firstParagraphText(handle.story(story.id))).toBe(original);

    const preview = recorder.op('previewProposal', () =>
      handle.previewProposal(rewrite.id)
    );
    expect(
      firstParagraphText(firstStory(preview.snapshot.slides[SLIDE].shapes))
    ).toBe('Proposed headline');
    expect(firstParagraphText(handle.story(story.id))).toBe(original);

    const ghost = recorder.op('layoutProposalSlide', () =>
      handle.layoutProposalSlide(rewrite.id, SLIDE)
    );
    expect(ghost.primitives.length).toBeGreaterThan(0);
    const diff = recorder.op('layoutProposalDiffSlide', () =>
      handle.layoutProposalDiffSlide(rewrite.id, SLIDE)
    );
    expect(diff.textChanges.length).toBeGreaterThan(0);
    expect(diff.textChanges.some((change) => change.kind === 'insertion')).toBe(
      true
    );

    expect(
      recorder.op('acceptProposal', () => handle.acceptProposal(rewrite.id))
        .applied
    ).toBe(true);
    expect(
      firstParagraphText(
        recorder.op('story:accepted', () => handle.story(story.id))
      )
    ).toBe('Proposed headline');
    expect(
      recorder.op('rejectProposal', () => handle.rejectProposal(recolour.id))
    ).toBe(true);
    expect(
      recorder.op('listProposals:afterReview', () => handle.listProposals())
    ).toEqual([]);
    expect(
      recorder
        .op('snapshot:afterReview', () => handle.snapshot())
        .slides[SLIDE].shapes.find((shape) => shape.id === target.id)
        ?.resolvedFillColor
    ).toBe(target.resolvedFillColor);

    profiledLayout(handle, recorder, 'layoutSlide:afterReview', SLIDE);
    const saved = recorder.op('save', () => handle.save());
    const reopened = recorder.op('reopen', () => open(saved));
    expect(
      recorder.op('searchText:afterReopen', () =>
        reopened.searchText('Proposed headline')
      ).length
    ).toBeGreaterThan(0);
  },
};

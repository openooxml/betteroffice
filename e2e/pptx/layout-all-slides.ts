import { expect } from 'bun:test';

import {
  editStages,
  firstStory,
  profiled,
  profiledLayout,
  runCount,
  storyText,
} from './context';
import type { PptxScenario } from './context';

export const layoutAllSlides: PptxScenario = {
  name: 'layout-all-slides',
  description:
    'Lays out every slide of the deck, edits the first story of each one and lays it out again, so cold and warm layout cost is measured per slide across the whole deck.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const deck = recorder.op('snapshot', () => handle.snapshot());
    expect(deck.slides.length).toBeGreaterThan(0);

    const widths = new Set<number>();
    for (let index = 0; index < deck.slides.length; index += 1) {
      const layout = profiledLayout(
        handle,
        recorder,
        'layoutSlide:cold',
        index
      );
      expect(layout.primitives.length).toBeGreaterThanOrEqual(0);
      widths.add(layout.width);
    }
    expect(widths.size).toBe(1);

    let edited = 0;
    for (let index = 0; index < deck.slides.length; index += 1) {
      const slide = deck.slides[index];
      let story;
      try {
        story = firstStory(slide.shapes);
      } catch {
        continue;
      }
      const before = storyText(story);
      const receipt = profiled(
        recorder,
        'insertText:perSlide',
        () => handle.insertTextProfiled(story.id, 0, `S${index} `),
        editStages
      );
      expect(receipt.storyId).toBe(story.id);
      expect(
        storyText(recorder.op('story:perSlide', () => handle.story(story.id)))
      ).toBe(`S${index} ${before}`);
      const warm = profiledLayout(handle, recorder, 'layoutSlide:warm', index);
      expect(runCount(warm)).toBeGreaterThan(0);
      edited += 1;
    }
    expect(edited).toBeGreaterThan(0);

    const saved = recorder.op('save', () => handle.save());
    const reopened = recorder.op('reopen', () => open(saved));
    expect(
      recorder.op('searchText:markers', () => reopened.searchText('S0 ')).length
    ).toBeGreaterThan(0);
    expect(
      recorder.op('snapshot:afterReopen', () => reopened.snapshot()).slides
        .length
    ).toBe(deck.slides.length);
  },
};

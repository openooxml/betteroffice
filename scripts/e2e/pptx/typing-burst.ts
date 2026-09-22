import { expect } from 'bun:test';

import { editStages, firstStory, laidOutText, profiled, profiledLayout, storyText, textBoxOf } from './context';
import type { PptxScenario } from './context';

const SLIDE = 0;
const BURST = 'Regional roll-up, provisional figures';

export const typingBurst: PptxScenario = {
  name: 'typing-burst',
  description: 'Types thirty-six characters one keystroke at a time into a live story, relaying out the slide after each, then deletes the burst in chunks; the per-keystroke and per-layout distributions are the measurement.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const slide = recorder.op('snapshot', () => handle.snapshot()).slides[SLIDE];
    const story = firstStory(slide.shapes);
    const original = storyText(story);

    for (let index = 0; index < BURST.length; index += 1) {
      const receipt = profiled(recorder, 'insertText:keystroke', () => handle.insertTextProfiled(story.id, index, BURST[index]), editStages);
      expect(receipt).toMatchObject({ storyId: story.id, start: index, end: index + 1 });
      profiledLayout(handle, recorder, 'layoutSlide:keystroke', SLIDE);
    }
    expect(storyText(recorder.op('story:afterBurst', () => handle.story(story.id)))).toBe(BURST + original);
    expect(laidOutText(textBoxOf(profiledLayout(handle, recorder, 'layoutSlide:settled', SLIDE), story.id))).toContain(BURST.slice(0, 8));

    for (let remaining = BURST.length; remaining > 0; remaining -= 10) {
      const start = Math.max(0, remaining - 10);
      const receipt = profiled(recorder, 'deleteText:chunk', () => handle.deleteTextProfiled(story.id, start, remaining), editStages);
      expect(receipt.text).toBe(BURST.slice(start, remaining));
      expect(storyText(recorder.op('story:afterChunk', () => handle.story(story.id)))).toBe(BURST.slice(0, start) + original);
    }
    expect(storyText(handle.story(story.id))).toBe(original);
  },
};

import { expect } from 'bun:test';

import { FONT_FAMILY, firstParagraphText, firstStory, profiledLayout, textBoxOf } from './context';
import type { PptxScenario } from './context';
import type { SlideDisplayList } from '../../../packages/pptx/src/types';

const SLIDE = 0;

/** The laid-out runs of the story's first paragraph, by story offset. */
function styledRuns(layout: SlideDisplayList, storyId: string, end: number) {
  return textBoxOf(layout, storyId)
    .lines.filter((line) => line.start < end)
    .flatMap((line) => line.runs)
    .filter((run) => run.start < end);
}

export const formattingSweep: PptxScenario = {
  name: 'formatting-sweep',
  description: 'Sweeps bold, italic, size, colour and alignment over every run of a story, checks the laid-out runs carry the new styling, then undoes the whole sweep and checks the layout returns to what it was.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const slide = recorder.op('snapshot', () => handle.snapshot()).slides[SLIDE];
    const story = firstStory(slide.shapes);
    const text = firstParagraphText(story);
    expect(text.length).toBeGreaterThan(0);

    const before = profiledLayout(handle, recorder, 'layoutSlide:before', SLIDE);
    const swatch = (run: { text: string; bold?: boolean; italic?: boolean; color?: string }) => ({ text: run.text, bold: run.bold, italic: run.italic, color: run.color });
    const originalRuns = styledRuns(before, story.id, text.length).map(swatch);
    const originalSize = styledRuns(before, story.id, text.length)[0].fontSizePx;
    expect(originalRuns.length).toBeGreaterThan(0);

    const steps = [
      { op: 'formatText:bold', patch: { bold: true } },
      { op: 'formatText:italic', patch: { italic: true } },
      { op: 'formatText:size', patch: { fontSizePt: 27 } },
      { op: 'formatText:color', patch: { color: '#118844' } },
      { op: 'formatText:family', patch: { fontFamily: FONT_FAMILY } },
    ];
    for (const step of steps) {
      expect(recorder.op(step.op, () => handle.formatText(story.id, 0, text.length, step.patch)).storyId).toBe(story.id);
    }
    expect(recorder.op('setParagraphAlignment', () => handle.setParagraphAlignment(story.id, 0, text.length, 'ctr')).storyId).toBe(story.id);

    const after = profiledLayout(handle, recorder, 'layoutSlide:formatted', SLIDE);
    const swept = styledRuns(after, story.id, text.length);
    expect(swept.length).toBeGreaterThan(0);
    expect(swept.every((run) => run.bold === true && run.italic === true)).toBe(true);
    expect(swept.every((run) => run.color?.toLowerCase() === '#118844')).toBe(true);
    expect(swept.map((run) => run.text).join('')).toBe(text);
    expect(swept.every((run) => run.fontSizePx !== originalSize)).toBe(true);

    for (let step = 0; step < steps.length + 1; step += 1) {
      expect(recorder.op('undo:sweep', () => handle.undo()).applied).toBe(true);
    }
    const restored = profiledLayout(handle, recorder, 'layoutSlide:restored', SLIDE);
    expect(styledRuns(restored, story.id, text.length).map(swatch)).toEqual(originalRuns);
    expect(styledRuns(restored, story.id, text.length)[0].fontSizePx).toBe(originalSize);
    expect(recorder.op('canUndo', () => handle.canUndo())).toBe(false);
  },
};

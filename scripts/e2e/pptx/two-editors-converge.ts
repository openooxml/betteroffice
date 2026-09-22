import { expect } from 'bun:test';

import { editStages, exchange, fingerprint, firstStory, profiled, replicas, storyText } from './context';
import type { PptxScenario } from './context';

const SLIDE = 0;
const BOX = { x: 914_400, y: 3_200_400, width: 2_743_200, height: 914_400 };

export const twoEditorsConverge: PptxScenario = {
  name: 'two-editors-converge',
  description: 'Two replicas type into different stories, then into the same position of one story, and one appends a slide while the other adds a box; after the updates are delivered both ways the deck fingerprints and the laid-out text must match.',
  participants: ['web:a', 'web:b'],
  run(ctx) {
    const [a, b] = replicas(ctx, ['web:a', 'web:b']);
    const deck = a.timer.op('snapshot', () => a.handle.snapshot());
    const slide = deck.slides[SLIDE];
    const story = firstStory(slide.shapes);
    const original = storyText(story);

    const first = profiled(a.timer, 'insertText:a', () => a.handle.insertTextProfiled(story.id, 0, 'A: '), editStages);
    expect(first.start).toBe(0);
    const other = deck.slides[SLIDE].shapes.find((shape) => shape.id !== slide.shapes[0].id && shape.textStories.length > 0);
    const secondStory = other ? other.textStories[0].id : story.id;
    profiled(b.timer, 'insertText:b', () => b.handle.insertTextProfiled(secondStory, 0, 'B: '), editStages);

    expect(exchange([a, b])).toBeGreaterThan(0);
    expect(fingerprint(a.timer.op('snapshot:merged', () => a.handle.snapshot()))).toBe(fingerprint(b.timer.op('snapshot:merged', () => b.handle.snapshot())));
    expect(storyText(a.handle.story(story.id))).toContain('A: ');
    expect(storyText(b.handle.story(story.id))).toContain('A: ');

    profiled(a.timer, 'insertText:sameSpot', () => a.handle.insertTextProfiled(story.id, 0, 'a'), editStages);
    profiled(b.timer, 'insertText:sameSpot', () => b.handle.insertTextProfiled(story.id, 0, 'b'), editStages);
    exchange([a, b], 'applyUpdate:conflict');
    const merged = storyText(a.handle.story(story.id));
    expect(storyText(b.handle.story(story.id))).toBe(merged);
    expect(merged).toContain('a');
    expect(merged).toContain('b');
    expect(merged.endsWith(original)).toBe(true);

    const appended = profiled(a.timer, 'insertSlide:a', () => a.handle.insertSlideProfiled(deck.slides.length), editStages);
    const added = profiled(
      b.timer,
      'addTextBox:b',
      () => b.handle.addTextBoxProfiled(slide.id, { name: 'from B', rect: BOX, text: 'Added by B', style: { fontSizePt: 16 } }),
      editStages
    );
    exchange([a, b], 'applyUpdate:structural');

    const left = a.timer.op('snapshot:final', () => a.handle.snapshot());
    const right = b.timer.op('snapshot:final', () => b.handle.snapshot());
    expect(fingerprint(left)).toBe(fingerprint(right));
    expect(left.slides.map((entry) => entry.id)).toContain(appended.slideId);
    expect(right.slides[SLIDE].shapes.map((shape) => shape.id)).toContain(added.shapeId);
    expect([...a.handle.encodeStateVector()].join()).toBe([...b.handle.encodeStateVector()].join());

    const catchUp = a.timer.op('encodeDiff:inSync', () => a.handle.encodeDiff(b.handle.encodeStateVector()));
    expect(catchUp.byteLength).toBeLessThan(a.handle.encodeStateAsUpdate().byteLength);
  },
};

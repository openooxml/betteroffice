import { expect } from 'bun:test';

import { editStages, profiled, profiledLayout } from './context';
import type { PptxScenario } from './context';

export const slideReorderAndDelete: PptxScenario = {
  name: 'slide-reorder-and-delete',
  description:
    'Inserts three slides, moves them around the deck and deletes one, checking the slide id order after every step and that undo and redo walk the order back and forward exactly.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const order = () =>
      recorder
        .op('snapshot:order', () => handle.snapshot())
        .slides.map((slide) => slide.id);
    const original = order();
    const base = original.length;

    const inserted: string[] = [];
    for (let index = 0; index < 3; index += 1) {
      const receipt = profiled(
        recorder,
        'insertSlide',
        () => handle.insertSlideProfiled(base + index),
        editStages
      );
      inserted.push(receipt.slideId);
    }
    const appended = order();
    expect(appended).toEqual([...original, ...inserted]);

    expect(
      recorder.op('moveSlide:toFront', () => handle.moveSlide(inserted[2], 0))
    ).toMatchObject({ slideId: inserted[2], toIndex: 0 });
    expect(order()).toEqual([
      inserted[2],
      ...original,
      inserted[0],
      inserted[1],
    ]);
    profiledLayout(handle, recorder, 'layoutSlide:moved', 0);

    expect(
      recorder.op('moveSlide:toMiddle', () => handle.moveSlide(inserted[0], 1))
        .toIndex
    ).toBe(1);
    const shuffled = order();
    expect(shuffled).toEqual([
      inserted[2],
      inserted[0],
      ...original,
      inserted[1],
    ]);

    expect(
      recorder.op('deleteSlide', () => handle.deleteSlide(inserted[1])).slideId
    ).toBe(inserted[1]);
    const trimmed = order();
    expect(trimmed).toEqual([inserted[2], inserted[0], ...original]);
    expect(trimmed).not.toContain(inserted[1]);

    expect(recorder.op('undo:delete', () => handle.undo()).applied).toBe(true);
    expect(order()).toEqual(shuffled);
    expect(recorder.op('undo:moveMiddle', () => handle.undo()).applied).toBe(
      true
    );
    expect(order()).toEqual([
      inserted[2],
      ...original,
      inserted[0],
      inserted[1],
    ]);
    expect(recorder.op('undo:moveFront', () => handle.undo()).applied).toBe(
      true
    );
    expect(order()).toEqual(appended);

    expect(recorder.op('redo:moveFront', () => handle.redo()).applied).toBe(
      true
    );
    expect(recorder.op('redo:moveMiddle', () => handle.redo()).applied).toBe(
      true
    );
    expect(order()).toEqual(shuffled);
    expect(recorder.op('redo:delete', () => handle.redo()).applied).toBe(true);
    expect(order()).toEqual(trimmed);
    profiledLayout(handle, recorder, 'layoutSlide:afterRedo', 0);
  },
};

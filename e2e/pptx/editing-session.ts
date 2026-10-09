import { expect } from 'bun:test';

import {
  FONT_FAMILY,
  editStages,
  firstStory,
  historyStages,
  laidOutText,
  profiled,
  profiledLayout,
  runCount,
  storyText,
  textBoxOf,
} from './context';
import type { PptxScenario } from './context';

const SLIDE = 0;
const MARKER = 'E2E ';
const NOTE = 'Added by the e2e run';
/** EMU, inside every slide of the corpus. */
const BOX = { x: 914_400, y: 914_400, width: 3_657_600, height: 914_400 };
const MOVED = { x: 1_828_800, y: 2_743_200 };

export const editingSession: PptxScenario = {
  name: 'editing-session',
  description:
    'One editor inserts text, adds and moves a text box, appends a slide, deletes text and undoes it, relaying out after each change, then saves and reopens.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const deck = recorder.op('snapshot', () => handle.snapshot());
    expect(deck.slides.length).toBeGreaterThan(0);
    const slide = deck.slides[SLIDE];
    const story = firstStory(slide.shapes);
    const original = storyText(story);

    const initial = profiledLayout(
      handle,
      recorder,
      'layoutSlide:initial',
      SLIDE
    );
    expect(initial.primitives.length).toBeGreaterThan(0);
    expect(runCount(initial)).toBeGreaterThan(0);
    expect(laidOutText(textBoxOf(initial, story.id)).length).toBeGreaterThan(0);

    const inserted = profiled(
      recorder,
      'insertText',
      () => handle.insertTextProfiled(story.id, 0, MARKER),
      editStages
    );
    expect(inserted).toMatchObject({
      storyId: story.id,
      start: 0,
      end: MARKER.length,
    });
    expect(
      storyText(
        recorder.op('story:afterInsertText', () => handle.story(story.id))
      )
    ).toBe(MARKER + original);
    const afterInsert = profiledLayout(
      handle,
      recorder,
      'layoutSlide:afterInsertText',
      SLIDE
    );
    expect(laidOutText(textBoxOf(afterInsert, story.id))).toContain(
      MARKER.trim()
    );

    const added = profiled(
      recorder,
      'addTextBox',
      () =>
        handle.addTextBoxProfiled(slide.id, {
          name: 'E2E note',
          rect: BOX,
          text: NOTE,
          style: { fontSizePt: 18, fontFamily: FONT_FAMILY },
        }),
      editStages
    );
    expect(added.slideId).toBe(slide.id);
    const withBox = recorder.op('snapshot:afterAddTextBox', () =>
      handle.snapshot()
    ).slides[SLIDE];
    expect(withBox.shapes.length).toBe(slide.shapes.length + 1);
    const box = withBox.shapes.find((shape) => shape.id === added.shapeId);
    expect(box).toMatchObject(BOX);
    const boxStory = box!.textStories[0].id;
    const afterAdd = profiledLayout(
      handle,
      recorder,
      'layoutSlide:afterAddTextBox',
      SLIDE
    );
    const boxBefore = textBoxOf(afterAdd, boxStory);
    expect(laidOutText(boxBefore)).toContain(NOTE.split(' ')[0]);

    const moved = profiled(
      recorder,
      'moveShape',
      () => handle.moveShapeProfiled(slide.id, added.shapeId, MOVED.x, MOVED.y),
      editStages
    );
    expect(moved.before).toMatchObject({ x: BOX.x, y: BOX.y });
    expect(moved.after).toMatchObject({
      ...MOVED,
      width: BOX.width,
      height: BOX.height,
    });
    const afterMove = profiledLayout(
      handle,
      recorder,
      'layoutSlide:afterMoveShape',
      SLIDE
    );
    const boxAfter = textBoxOf(afterMove, boxStory);
    expect(boxAfter.x).toBeGreaterThan(boxBefore.x);
    expect(boxAfter.y).toBeGreaterThan(boxBefore.y);

    const count = deck.slides.length;
    const appended = profiled(
      recorder,
      'insertSlide',
      () => handle.insertSlideProfiled(count),
      editStages
    );
    expect(appended.toIndex).toBe(count);
    const grown = recorder.op('snapshot:afterInsertSlide', () =>
      handle.snapshot()
    );
    expect(grown.slides.length).toBe(count + 1);
    expect(grown.slides[count].id).toBe(appended.slideId);
    const blank = profiledLayout(
      handle,
      recorder,
      'layoutSlide:insertedSlide',
      count
    );
    expect(blank).toMatchObject({
      width: initial.width,
      height: initial.height,
    });

    const deleted = profiled(
      recorder,
      'deleteText',
      () => handle.deleteTextProfiled(story.id, 0, MARKER.length),
      editStages
    );
    expect(deleted.text).toBe(MARKER);
    expect(
      storyText(
        recorder.op('story:afterDeleteText', () => handle.story(story.id))
      )
    ).toBe(original);

    const undone = profiled(
      recorder,
      'undo:deleteText',
      () => handle.undoProfiled(),
      historyStages
    );
    expect(undone.applied).toBe(true);
    expect(undone.snapshot.slides.length).toBe(count + 1);
    expect(
      storyText(recorder.op('story:afterUndo', () => handle.story(story.id)))
    ).toBe(MARKER + original);

    const saved = recorder.op('save', () => handle.save());
    expect(saved.byteLength).toBeGreaterThan(0);

    const reopened = recorder.op('reopen', () => open(saved));
    const again = recorder.op('snapshot:afterReopen', () =>
      reopened.snapshot()
    );
    expect(again.slides.length).toBe(count + 1);
    expect(again.slides[SLIDE].shapes.length).toBe(slide.shapes.length + 1);
    const matches = recorder.op('searchText:afterReopen', () =>
      reopened.searchText(MARKER.trim())
    );
    expect(
      matches.some((match) => match.slideIndex === SLIDE && match.start === 0)
    ).toBe(true);
  },
};

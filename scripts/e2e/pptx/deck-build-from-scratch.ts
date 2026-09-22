import { expect } from 'bun:test';

import {
  FONT_FAMILY,
  editStages,
  laidOutText,
  profiled,
  profiledLayout,
  shapeText,
  storyText,
  textBoxOf,
} from './context';
import type { PptxScenario } from './context';

const SLIDES = 8;
const TITLE = { x: 685_800, y: 457_200, width: 7_772_400, height: 1_143_000 };
const BODY = { x: 685_800, y: 1_828_800, width: 5_486_400, height: 2_743_200 };
const BADGE = {
  x: 6_400_800,
  y: 1_828_800,
  width: 1_600_200,
  height: 1_600_200,
};

export const deckBuildFromScratch: PptxScenario = {
  name: 'deck-build-from-scratch',
  description:
    'Appends eight slides and builds each one out of a title box, a body box and a preset shape, formats the title, moves and resizes the badge, lays every new slide out and checks it all again after a save and reopen.',
  participants: ['web'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const before = recorder.op('snapshot', () => handle.snapshot());
    const base = before.slides.length;
    const titles: string[] = [];

    for (let index = 0; index < SLIDES; index += 1) {
      const title = `Section ${index + 1}`;
      titles.push(title);
      const appended = profiled(
        recorder,
        'insertSlide',
        () => handle.insertSlideProfiled(base + index),
        editStages
      );
      expect(appended.toIndex).toBe(base + index);

      const titleBox = profiled(
        recorder,
        'addTextBox:title',
        () =>
          handle.addTextBoxProfiled(appended.slideId, {
            name: `title ${index}`,
            rect: TITLE,
            text: title,
            style: { fontSizePt: 32, fontFamily: FONT_FAMILY, bold: true },
          }),
        editStages
      );
      const bodyBox = recorder.op('addTextBox:body', () =>
        handle.addTextBox(appended.slideId, {
          name: `body ${index}`,
          rect: BODY,
          text: `Body copy for ${title}`,
          style: { fontSizePt: 18, fontFamily: FONT_FAMILY },
        })
      );
      const badge = recorder.op('addShape:badge', () =>
        handle.addShape(appended.slideId, {
          name: `badge ${index}`,
          geometry: 'ellipse',
          rect: BADGE,
          fill: '#3366cc',
        })
      );

      const slide = recorder.op('snapshot:slide', () => handle.snapshot())
        .slides[base + index];
      expect(slide.shapes.map((shape) => shape.id)).toEqual([
        titleBox.shapeId,
        bodyBox.shapeId,
        badge.shapeId,
      ]);
      expect(shapeText(slide.shapes[0])).toBe(title);
      expect(slide.shapes[2]).toMatchObject({ geometry: 'ellipse', ...BADGE });

      const titleStory = slide.shapes[0].textStories[0].id;
      expect(
        recorder.op('formatText:title', () =>
          handle.formatText(titleStory, 0, title.length, {
            italic: true,
            color: '#cc0033',
          })
        ).storyId
      ).toBe(titleStory);
      expect(
        recorder.op('resizeShape:badge', () =>
          handle.resizeShape(
            appended.slideId,
            badge.shapeId,
            1_200_150,
            1_200_150
          )
        ).after
      ).toMatchObject({ width: 1_200_150, height: 1_200_150 });
      expect(
        profiled(
          recorder,
          'moveShape:badge',
          () =>
            handle.moveShapeProfiled(
              appended.slideId,
              badge.shapeId,
              BADGE.x,
              BADGE.y + 457_200
            ),
          editStages
        ).after.y
      ).toBe(BADGE.y + 457_200);

      const layout = profiledLayout(
        handle,
        recorder,
        'layoutSlide:built',
        base + index
      );
      expect(laidOutText(textBoxOf(layout, titleStory))).toBe(title);
    }

    const built = recorder.op('snapshot:built', () => handle.snapshot());
    expect(built.slides.length).toBe(base + SLIDES);

    const saved = recorder.op('save', () => handle.save());
    const reopened = recorder.op('reopen', () => open(saved));
    const again = recorder.op('snapshot:afterReopen', () =>
      reopened.snapshot()
    );
    expect(again.slides.length).toBe(base + SLIDES);
    for (let index = 0; index < SLIDES; index += 1) {
      const slide = again.slides[base + index];
      expect(slide.shapes.length).toBe(3);
      expect(storyText(slide.shapes[0].textStories[0])).toBe(titles[index]);
      expect(shapeText(slide.shapes[1])).toContain(titles[index]);
    }
  },
};

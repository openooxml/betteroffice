import { expect } from 'bun:test';

import { STORY, displayFrame, frameText, paragraphText, regionLayout, typing, typingTarget } from './context';
import type { DocxScenario } from './context';

const PHRASE = 'Inserted paragraph for pagination pressure, long enough to reflow a line. ';
const SPLITS = 10;

export const paginationPressure: DocxScenario = {
  name: 'pagination-pressure',
  description: 'Types the same phrase at the start, the middle and the end of the body and then splits the last paragraph ten times, tracking the page count the engine reports after every change.',
  participants: ['web'],
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;
    const paragraphs = session.paragraphs(STORY);
    const spans = new Map(session.paragraphSpans(STORY).map((span) => [span.paraId, span.length]));
    const typable = paragraphs.filter((paragraph) => /\S/.test(paragraph.text));
    expect(typable.length).toBeGreaterThan(0);

    const initial = regionLayout(editor, recorder, 'layoutDocumentWithRegions:initial');
    const pagesBefore = initial.layout.pages.length;
    expect(pagesBefore).toBeGreaterThan(0);
    let frame = displayFrame(session, recorder, 'displayListFrame:initial', null);
    expect(frame.pages.length).toBe(pagesBefore);

    const spots = [typable[0], typable[Math.floor(typable.length / 2)], typable[typable.length - 1]];
    const labels = ['start', 'middle', 'end'];
    spots.forEach((paragraph, index) => {
      const offset = spans.get(paragraph.paraId) ?? paragraph.text.length;
      recorder.op(`setSelection:${labels[index]}`, () => session.setSelection({ story: STORY, paraId: paragraph.paraId, offset }));
      frame = typing(session, recorder, `applyInput:${labels[index]}`, PHRASE, frame);
      expect(paragraphText(session, paragraph.paraId).endsWith(PHRASE)).toBe(true);
      const laid = regionLayout(editor, recorder, `layoutDocumentWithRegions:${labels[index]}`);
      expect(laid.layout.pages.length).toBeGreaterThanOrEqual(pagesBefore);
    });
    expect(frameText(frame)).toContain('pagination pressure');

    const tail = spots[2];
    let offset = (spans.get(tail.paraId) ?? 0) + PHRASE.length;
    let paraId = tail.paraId;
    for (let index = 0; index < SPLITS; index += 1) {
      const split = recorder.op('splitParagraph:tail', () => session.splitParagraph({ story: STORY, paraId, offset }));
      paraId = split.secondParaId;
      offset = 0;
      frame = typing(session, recorder, 'applyInput:tail', PHRASE, frame);
    }
    const grown = regionLayout(editor, recorder, 'layoutDocumentWithRegions:afterSplits');
    expect(grown.layout.pages.length).toBeGreaterThanOrEqual(pagesBefore);
    expect(session.paragraphs(STORY).length).toBe(paragraphs.length + SPLITS);
    frame = displayFrame(session, recorder, 'displayListFrame:afterSplits', frame);
    expect(frame.pages.length).toBe(grown.layout.pages.length);
  },
};

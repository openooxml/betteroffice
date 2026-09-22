import { expect } from 'bun:test';

import { STORY, displayFrame, frameText, regionLayout } from './context';
import type { DocxScenario } from './context';

const REPLACEMENT = 'Zenith';
const LIMIT = 30;

export const searchAndReplaceSweep: DocxScenario = {
  name: 'search-and-replace-sweep',
  description: 'Finds the most frequent word in the document and replaces up to thirty occurrences one at a time, re-searching after each so the sweep follows the moving offsets, then checks the old word is gone and the new one is everywhere.',
  participants: ['web'],
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;
    const query = frequentWord(session.paragraphs(STORY).map((paragraph) => paragraph.text).join(' '));
    const found = recorder.op('searchText:initial', () => session.searchText(query));
    expect(found.length).toBeGreaterThan(0);
    const total = Math.min(found.length, LIMIT);

    for (let index = 0; index < total; index += 1) {
      const matches = recorder.op('searchText:next', () => session.searchText(query));
      expect(matches.length).toBe(found.length - index);
      const match = matches[0];
      const range = { story: match.story, start: { paraId: match.paraId, offset: match.start }, end: { paraId: match.paraId, offset: match.end } };
      recorder.op('replaceRange', () => session.replaceRange(range, REPLACEMENT));
      const paragraph = session.paragraphs(match.story).find((entry) => entry.paraId === match.paraId);
      expect(paragraph?.text.includes(REPLACEMENT)).toBe(true);
    }

    const remaining = recorder.op('searchText:afterSweep', () => session.searchText(query));
    expect(remaining.length).toBe(found.length - total);
    const replaced = recorder.op('searchText:replacement', () => session.searchText(REPLACEMENT));
    expect(replaced.length).toBeGreaterThanOrEqual(total);

    regionLayout(editor, recorder, 'layoutDocumentWithRegions:afterSweep');
    expect(frameText(displayFrame(session, recorder, 'displayListFrame:afterSweep', null))).toContain(REPLACEMENT);
  },
};

/** The most common word of at least five letters, so the sweep has many hits. */
function frequentWord(text: string): string {
  const counts = new Map<string, number>();
  for (const word of text.match(/[A-Za-z]{5,}/g) ?? []) counts.set(word, (counts.get(word) ?? 0) + 1);
  const best = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0];
  if (!best) throw new Error('the document has no word to sweep');
  return best[0];
}

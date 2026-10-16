import { expect } from 'bun:test';

import {
  STORY,
  paragraphText,
  regionLayout,
  save,
  typingTarget,
} from './context';
import type { DocxScenario } from './context';

const CYCLES = 10;

export const saveLoadCycles: DocxScenario = {
  name: 'save-load-cycles',
  description:
    'Ten rounds of type, project, repack and reopen in a fresh session, checking every earlier marker survives and file growth stays bounded while recording save/open timings.',
  participants: ['web'],
  async run({ recorder, open }) {
    let editor = await recorder.loadAsync(() => open());
    let previousBytes = 0;

    for (let cycle = 0; cycle < CYCLES; cycle += 1) {
      const { session } = editor;
      const target = typingTarget(
        session,
        session.paragraphs(STORY),
        session.paragraphSpans(STORY)
      );
      const marker = `[cycle ${cycle}]`;
      recorder.op('insertText:cycle', () =>
        session.insertText(
          { story: STORY, paraId: target.paraId, offset: target.end },
          marker
        )
      );
      expect(paragraphText(session, target.paraId).endsWith(marker)).toBe(true);
      regionLayout(editor, recorder, 'layoutDocumentWithRegions:cycle');

      const saved = await save(editor, recorder);
      if (previousBytes > 0)
        expect(Math.abs(saved.byteLength - previousBytes)).toBeLessThan(
          previousBytes * 0.25
        );
      previousBytes = saved.byteLength;

      editor = await recorder.opAsync('reopen:cycle', () => open(saved));
      const text = recorder
        .op('paragraphs:cycle', () => editor.session.paragraphs(STORY))
        .map((paragraph) => paragraph.text)
        .join('\n');
      for (let earlier = 0; earlier <= cycle; earlier += 1)
        expect(text).toContain(`[cycle ${earlier}]`);
    }

    const found = recorder.op('searchText:allCycles', () =>
      editor.session.searchText('[cycle ')
    );
    expect(found.length).toBeGreaterThanOrEqual(CYCLES);
  },
};

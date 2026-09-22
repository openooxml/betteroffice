import { expect } from 'bun:test';

import {
  STORY,
  UNDO_STEP_GAP_MS,
  displayFrame,
  frameText,
  matchRange,
  regionLayout,
  typingTarget,
} from './context';
import type { DocxScenario } from './context';
import type { YrsStorySegment } from '../../../packages/docx/src/yrs';

const WORD = 'Meridian';
const LINK = 'https://openooxml.dev';

export const formattingAndStyles: DocxScenario = {
  name: 'formatting-and-styles',
  description:
    'Marks, set-valued run formatting, a paragraph style, a hyperlink and a clear pass over one typed word, each read back through selectionContext and the rendered frame, then undone step by step.',
  participants: ['web'],
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;
    const target = typingTarget(
      session,
      session.paragraphs(STORY),
      session.paragraphSpans(STORY)
    );
    const at = { story: STORY, paraId: target.paraId, offset: target.end };
    recorder.op('insertText:word', () => session.insertText(at, WORD));
    const range = matchRange(
      recorder.op('searchText:word', () => session.searchText(WORD)),
      target.paraId,
      WORD.length
    );
    const plain = recorder.op('selectionContext:plain', () =>
      session.selectionContext(range)
    );
    expect(plain.bold).toBe(false);

    await Bun.sleep(UNDO_STEP_GAP_MS);
    recorder.op('toggleMark:bold', () =>
      session.toggleMark(range, { type: 'bold' })
    );
    recorder.op('toggleMark:italic', () =>
      session.toggleMark(range, { type: 'italic' })
    );
    const marked = recorder.op('selectionContext:marked', () =>
      session.selectionContext(range)
    );
    expect(marked.bold).toBe(true);
    expect(marked.italic).toBe(!plain.italic);

    await Bun.sleep(UNDO_STEP_GAP_MS);
    recorder.op('formatRange', () =>
      session.formatRange(range, { fontSize: 18, color: { rgb: 'CC2244' } })
    );
    const formatted = recorder.op('selectionContext:formatted', () =>
      session.selectionContext(range)
    );
    expect(formatted.fontSize).toBe(36);
    expect(formatted.color).toBe('CC2244');

    await Bun.sleep(UNDO_STEP_GAP_MS);
    recorder.op('setHyperlink', () =>
      session.setHyperlink(range, { href: LINK })
    );
    expect(
      linkedText(
        recorder.op('storySegments:linked', () => session.storySegments(STORY))
      )
    ).toContain(WORD);

    await Bun.sleep(UNDO_STEP_GAP_MS);
    const whole = {
      story: STORY,
      start: { paraId: target.paraId, offset: 0 },
      end: { paraId: target.paraId, offset: target.end + WORD.length },
    };
    recorder.op('applyParagraphStyle', () =>
      session.applyParagraphStyle(whole, 'Heading1')
    );
    recorder.op('setParagraphAttrs', () =>
      session.setParagraphAttrs(whole, { alignment: 'center' })
    );

    const laid = regionLayout(
      editor,
      recorder,
      'layoutDocumentWithRegions:styled'
    );
    expect(laid.layout.pages.length).toBeGreaterThan(0);
    expect(
      frameText(
        displayFrame(session, recorder, 'displayListFrame:styled', null)
      )
    ).toContain(WORD);

    await Bun.sleep(UNDO_STEP_GAP_MS);
    recorder.op('clearFormatting', () => session.clearFormatting(range));
    const cleared = recorder.op('selectionContext:cleared', () =>
      session.selectionContext(range)
    );
    expect(cleared.bold).toBe(false);
    expect(cleared.color).not.toBe('CC2244');

    expect(recorder.op('undo:clear', () => session.undo())).toBe(true);
    expect(
      recorder.op('selectionContext:afterUndo', () =>
        session.selectionContext(range)
      ).bold
    ).toBe(true);
    expect(recorder.op('undo:paragraph', () => session.undo())).toBe(true);
    expect(recorder.op('undo:hyperlink', () => session.undo())).toBe(true);
    expect(
      linkedText(
        recorder.op('storySegments:unlinked', () =>
          session.storySegments(STORY)
        )
      )
    ).not.toContain(WORD);
    expect(recorder.op('undo:formatRange', () => session.undo())).toBe(true);
    expect(
      recorder.op('selectionContext:unformatted', () =>
        session.selectionContext(range)
      ).fontSize
    ).not.toBe(36);
    expect(recorder.op('undo:marks', () => session.undo())).toBe(true);
    const unmarked = recorder.op('selectionContext:unmarked', () =>
      session.selectionContext(range)
    );
    expect(unmarked.bold).toBe(false);
    expect(unmarked.italic).toBe(plain.italic);
  },
};

/** The text of every run the story marks as a hyperlink. */
function linkedText(segments: YrsStorySegment[]): string {
  return segments
    .filter(
      (segment) =>
        segment.kind === 'text' &&
        JSON.stringify(segment.attributes).includes(LINK)
    )
    .map((segment) => (segment.kind === 'text' ? segment.text : ''))
    .join('');
}

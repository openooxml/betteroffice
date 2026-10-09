import { expect } from 'bun:test';

import {
  STORY,
  UNDO_STEP_GAP_MS,
  blocks,
  displayFrame,
  frameText,
  matchRange,
  paragraphText,
  regionLayout,
  save,
  typing,
  typingTarget,
} from './context';
import type { DocxScenario } from './context';
import { yrsToDocument } from '../../packages/docx/src/yrs';

const WORD = 'Zephyr';
const PHRASE = 'quartz lantern drifts over the harbor';
const TYPED = `${WORD} ${PHRASE}`;

export const editingSession: DocxScenario = {
  name: 'editing-session',
  description:
    'One editor types a word, a space and a phrase through the resident layout path, splits the paragraph, inserts a table, bolds the typed word, undoes and redoes it, then saves and reopens the document.',
  participants: ['web'],
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;

    const paragraphs = recorder.op('paragraphs:body', () =>
      session.paragraphs(STORY)
    );
    expect(paragraphs.length).toBeGreaterThan(0);
    const spans = recorder.op('paragraphSpans:body', () =>
      session.paragraphSpans(STORY)
    );
    const target = typingTarget(session, paragraphs, spans);

    const initial = regionLayout(
      editor,
      recorder,
      'layoutDocumentWithRegions:initial'
    );
    expect(initial.layout.pages.length).toBeGreaterThan(0);

    const caret = { story: STORY, paraId: target.paraId, offset: target.end };
    recorder.op('setSelection', () => session.setSelection(caret));
    expect(session.encodeSelection()?.story).toBe(STORY);

    let frame = displayFrame(
      session,
      recorder,
      'displayListFrame:initial',
      null
    );
    expect(frame.pages.length).toBe(initial.layout.pages.length);
    expect(frame.displayList.pages[0].primitives.length).toBeGreaterThan(0);

    const snapshot = recorder.op('residentCaretSnapshot', () =>
      session.residentCaretSnapshot()
    );
    expect(snapshot.frameEpoch).toBe(frame.frameEpoch);
    expect(snapshot.caretRect).not.toBeNull();

    const expectTyped = (text: string) => {
      expect(paragraphText(session, caret.paraId).endsWith(text)).toBe(true);
      expect(session.selection()?.head).toEqual({
        story: STORY,
        paraId: caret.paraId,
        offset: caret.offset + text.length,
      });
    };

    frame = typing(session, recorder, 'applyInput:word', WORD, frame);
    expectTyped(WORD);
    expect(frameText(frame).toLowerCase()).toContain(WORD.toLowerCase());

    frame = typing(session, recorder, 'applyInput:space', ' ', frame);
    expectTyped(`${WORD} `);

    frame = typing(session, recorder, 'applyInput:phrase', PHRASE, frame);
    expectTyped(TYPED);
    expect(frameText(frame).toLowerCase()).toContain('harbor');

    const split = recorder.op('splitParagraph', () =>
      session.splitParagraph(caret)
    );
    expect(split.firstParaId).toBe(target.paraId);
    expect(paragraphText(session, target.paraId)).toBe(target.text);
    const typedParaId = split.secondParaId;
    expect(paragraphText(session, typedParaId)).toBe(TYPED);

    const storiesBefore = session.storyIds().length;
    const table = recorder.op('insertTable', () =>
      session.insertTable(
        { story: STORY, paraId: typedParaId, offset: 0 },
        2,
        2
      )
    );
    expect(table).toMatchObject({ rows: 2, columns: 2, deletedTable: false });
    expect(table.createdStoryIds).toHaveLength(4);
    expect(session.storyIds()).toHaveLength(storiesBefore + 4);
    expect(paragraphText(session, typedParaId)).toBe(TYPED);

    const range = matchRange(
      recorder.op('searchText:typed', () => session.searchText(WORD)),
      typedParaId,
      WORD.length
    );
    expect(
      recorder.op('selectionContext', () => session.selectionContext(range))
        .bold
    ).toBe(false);
    await Bun.sleep(UNDO_STEP_GAP_MS);
    recorder.op('toggleMark:bold', () =>
      session.toggleMark(range, { type: 'bold' })
    );
    expect(session.selectionContext(range).bold).toBe(true);

    expect(recorder.op('undo:toggleMark', () => session.undo())).toBe(true);
    expect(session.selectionContext(range).bold).toBe(false);
    expect(paragraphText(session, typedParaId)).toBe(TYPED);
    expect(recorder.op('redo:toggleMark', () => session.redo())).toBe(true);
    expect(session.selectionContext(range).bold).toBe(true);

    const edited = regionLayout(
      editor,
      recorder,
      'layoutDocumentWithRegions:afterEdits'
    );
    expect(edited.layout.pages.length).toBeGreaterThanOrEqual(
      initial.layout.pages.length
    );
    frame = displayFrame(
      session,
      recorder,
      'displayListFrame:afterEdits',
      frame
    );
    expect(frame.pages.length).toBe(edited.layout.pages.length);
    expect(frameText(frame).toLowerCase()).toContain(WORD.toLowerCase());

    const materialized = session.materializeDocx()!;
    const projected = yrsToDocument(session, materialized);
    expect(blocks(projected, 'paragraph')).toBe(
      blocks(materialized, 'paragraph') + 1
    );
    expect(blocks(projected, 'table')).toBe(blocks(materialized, 'table') + 1);
    const saved = await save(editor, recorder);

    const reopened = await recorder.opAsync('reopen', () => open(saved));
    expect(reopened.session.storyIds()).toHaveLength(session.storyIds().length);
    const persisted = recorder
      .op('paragraphs:afterReopen', () => reopened.session.paragraphs(STORY))
      .find((paragraph) => paragraph.text === TYPED);
    expect(persisted).toBeDefined();
    const found = recorder.op('searchText:afterReopen', () =>
      reopened.session.searchText(WORD)
    );
    expect(
      reopened.session.selectionContext(
        matchRange(found, persisted!.paraId, WORD.length)
      ).bold
    ).toBe(true);
  },
};

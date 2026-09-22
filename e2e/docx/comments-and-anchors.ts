import { expect } from 'bun:test';

import { STORY, matchRange, typingTarget } from './context';
import type { DocxScenario } from './context';

const WORD = 'Cornerstone';
const DATE = '2026-09-22T00:00:00Z';

export const commentsAndAnchors: DocxScenario = {
  name: 'comments-and-anchors',
  description:
    'Anchors two comments over typed text, checks the anchors resolve back to the ranges they were placed on, then edits around them and checks the anchors moved with the text rather than stayed put.',
  participants: ['web', 'reviewer'],
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;
    const target = typingTarget(
      session,
      session.paragraphs(STORY),
      session.paragraphSpans(STORY)
    );
    const at = { story: STORY, paraId: target.paraId, offset: target.end };
    recorder.op('insertText:word', () =>
      session.insertText(at, ` ${WORD} tail`)
    );
    const range = matchRange(
      recorder.op('searchText:word', () => session.searchText(WORD)),
      target.paraId,
      WORD.length
    );

    const first = recorder.op(
      'addComment:first',
      () =>
        session.addComment([range], 'Reviewer', DATE, {
          text: 'Is this the right term?',
        }),
      undefined,
      { actor: 'reviewer' }
    );
    expect(first.commentId).toBeTruthy();
    const second = recorder.op(
      'addComment:second',
      () =>
        session.addComment([range], 'Second reviewer', DATE, {
          text: 'Agreed',
        }),
      undefined,
      { actor: 'reviewer' }
    );
    expect(second.commentId).not.toBe(first.commentId);

    const anchors = recorder.op('resolveComment:first', () =>
      session.resolveComment(first.commentId)
    );
    expect(anchors.length).toBeGreaterThan(0);
    expect(anchors[0].story).toBe(STORY);
    expect(anchors[0].end - anchors[0].start).toBe(WORD.length);
    const before = anchors[0].start;

    recorder.op('insertText:ahead', () =>
      session.insertText(
        { story: STORY, paraId: target.paraId, offset: 0 },
        'PREFIX '
      )
    );
    const moved = recorder.op('resolveComment:afterInsert', () =>
      session.resolveComment(first.commentId)
    );
    expect(moved[0].start).toBe(before + 'PREFIX '.length);
    expect(moved[0].end - moved[0].start).toBe(WORD.length);

    const still = recorder.op('resolveComment:second', () =>
      session.resolveComment(second.commentId)
    );
    expect(still[0].start).toBe(moved[0].start);
    expect(() =>
      recorder.op('resolveComment:unknown', () =>
        session.resolveComment('no-such-comment')
      )
    ).toThrow('no-such-comment');
  },
};

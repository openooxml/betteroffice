import { expect } from 'bun:test';

import type { PptxScenario } from './context';

const SLIDE = 0;
const CREATED = '2026-09-22T00:00:00Z';

export const commentsThread: PptxScenario = {
  name: 'comments-thread',
  description: 'Opens a review thread on a slide, replies to it, resolves and reopens it, removes a second thread, and checks the comment snapshots and slide notes survive a save and reopen.',
  participants: ['web', 'reviewer'],
  run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const slide = recorder.op('snapshot', () => handle.snapshot()).slides[SLIDE];
    const existing = recorder.op('comments:initial', () => handle.comments()).length;
    expect(existing).toBe(0);
    expect(recorder.op('setCommentFlavor', () => handle.setCommentFlavor('modern'))).toBe('modern');

    const root = recorder.op('addComment:root', () => handle.addComment(slide.id, { author: 'Reviewer', initials: 'RV', text: 'Tighten this headline', created: CREATED, xEmu: 914_400, yEmu: 457_200 }), undefined, { actor: 'reviewer' });
    const second = recorder.op('addComment:second', () => handle.addComment(slide.id, { author: 'Reviewer', initials: 'RV', text: 'Check the figures', created: CREATED }), undefined, { actor: 'reviewer' });
    const opened = recorder.op('comments:afterAdd', () => handle.comments());
    expect(opened.length).toBe(existing + 2);
    expect(opened.find((comment) => comment.id === root.commentId)).toMatchObject({ slideId: slide.id, author: 'Reviewer', text: 'Tighten this headline', resolved: false });

    const reply = recorder.op('replyToComment', () => handle.replyToComment(root.commentId, { author: 'Author', initials: 'AU', text: 'Reworded', created: CREATED }));
    const threaded = recorder.op('comments:afterReply', () => handle.comments());
    expect(threaded.find((comment) => comment.id === reply.commentId)?.parentId).toBe(root.commentId);

    expect(recorder.op('setCommentStatus:resolve', () => handle.setCommentStatus(root.commentId, true)).commentId).toBe(root.commentId);
    expect(recorder.op('comments:resolved', () => handle.comments()).find((comment) => comment.id === root.commentId)?.resolved).toBe(true);
    expect(recorder.op('setCommentStatus:reopen', () => handle.setCommentStatus(root.commentId, false)).commentId).toBe(root.commentId);
    expect(recorder.op('comments:reopened', () => handle.comments()).find((comment) => comment.id === root.commentId)?.resolved).toBe(false);

    expect(recorder.op('removeComment', () => handle.removeComment(second.commentId)).commentId).toBe(second.commentId);
    expect(recorder.op('comments:afterRemove', () => handle.comments()).some((comment) => comment.id === second.commentId)).toBe(false);

    recorder.op('setSlideNotes', () => handle.setSlideNotes(slide.id, 'Reviewed in the e2e run'));

    const saved = recorder.op('save', () => handle.save());
    const reopened = recorder.op('reopen', () => open(saved));
    const persisted = recorder.op('comments:afterReopen', () => reopened.comments());
    expect(persisted.some((comment) => comment.text === 'Tighten this headline')).toBe(true);
    expect(persisted.some((comment) => comment.text === 'Reworded')).toBe(true);
    expect(persisted.some((comment) => comment.text === 'Check the figures')).toBe(false);
  },
};

import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsStoryRange } from './index';

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);

function range(story: string, paraId: string, start: number, end: number): YrsStoryRange {
  return { story, start: { paraId, offset: start }, end: { paraId, offset: end } };
}

describe('setCommentRanges', () => {
  it('moves anchors across paragraphs and stories with independent undo and redo', async () => {
    const session = await createYrsSession({ clientId: 80001 });
    try {
      const [paraId, secondParaId] = session.loadStories([
        {
          storyId: 'body',
          paragraphs: [{ text: 'first' }, { text: ' paragraph' }],
        },
      ]).body;
      const { paraId: headerId } = session.createStory('hf:rId1', 'header text');
      const { commentId } = session.addComment(
        [range('body', paraId, 0, 5)],
        'Ada',
        '2026-09-22',
        []
      );
      session.beginUndoCapture();
      const before = session.resolveComment(commentId);
      let updates = 0;
      session.onUpdate(() => { updates += 1; });
      session.setCommentRanges(commentId, [
        { story: 'body', start: { paraId, offset: 1 }, end: { paraId: secondParaId, offset: 4 } },
        range('hf:rId1', headerId, 0, 6),
      ]);
      const after = [
        { story: 'body', start: 1, end: 10 },
        { story: 'hf:rId1', start: 0, end: 6 },
      ];
      expect(session.resolveComment(commentId)).toEqual(after);
      expect(updates).toBe(1);
      for (let i = 0; i < 3; i += 1) {
        expect(session.undo()).toBe(true);
        expect(session.resolveComment(commentId)).toEqual(before);
        expect(session.historyStories()).toEqual(['body', 'hf:rId1']);
        expect(session.redo()).toBe(true);
        expect(session.resolveComment(commentId)).toEqual(after);
        expect(updates).toBe(3 + i * 2);
      }
    } finally {
      session.destroy();
    }
  });

  it('rejects all invalid ranges before changing document state or history', async () => {
    const session = await createYrsSession({ clientId: 80002 });
    try {
      const { paraId } = session.createStory('body', 'Achado preservado');
      const valid = range('body', paraId, 0, 6);
      const { commentId } = session.addComment([valid], 'Ada', '2026-09-22', []);
      session.beginUndoCapture();
      const state = session.encodeState();
      const invalid: YrsStoryRange[][] = [
        [],
        [range('body', paraId, 0, 0)],
        [range('body', paraId, 4, 2)],
        [range('missing', paraId, 0, 2)],
        [range('body', 'missing', 0, 2)],
        ...[-1, 0.5, NaN, Infinity, 0x100000000, Number.MAX_SAFE_INTEGER, 99].map((end) => [
          valid,
          range('body', paraId, 0, end),
        ]),
      ];
      for (const ranges of invalid) {
        expect(() => session.setCommentRanges(commentId, ranges)).toThrow();
        expect(session.encodeState()).toEqual(state);
        expect(session.canUndo()).toBe(false);
      }
      expect(() => session.setCommentRanges('missing', [valid])).toThrow();
      expect(session.encodeState()).toEqual(state);
    } finally {
      session.destroy();
    }
  });

  it('rejects empty reanchoring after all commented text is removed and preserves undo', async () => {
    const session = await createYrsSession({ clientId: 80007 });
    try {
      const original = 'Achado preservado. Conclusão antiga.';
      const { paraId } = session.createStory('body', original);
      const { commentId } = session.addComment(
        [range('body', paraId, 0, 17)], 'Ada', '2026-09-22', []
      );
      const before = session.resolveComment(commentId);
      session.beginUndoCapture();
      session.replaceRange(range('body', paraId, 0, original.length), '');
      const removed = session.encodeState();
      for (const ranges of [[], [range('body', paraId, 0, 0)], [range('body', paraId, 0, 17)]]) {
        expect(() => session.setCommentRanges(commentId, ranges)).toThrow();
        expect(session.encodeState()).toEqual(removed);
      }
      expect(session.paragraphs('body')[0].text).toBe('');
      expect(session.undo()).toBe(true);
      expect(session.paragraphs('body')[0].text).toBe(original);
      expect(session.resolveComment(commentId)).toEqual(before);
      expect(session.undo()).toBe(false);
      expect(session.redo()).toBe(true);
      expect(session.paragraphs('body')[0].text).toBe('');
    } finally {
      session.destroy();
    }
  });

  it('replicates new anchors without making remote changes locally undoable', async () => {
    const source = await createYrsSession({ clientId: 80003 });
    const peer = await createYrsSession({ clientId: 80004 });
    try {
      const { paraId } = source.createStory('body', 'first second');
      const { commentId } = source.addComment(
        [range('body', paraId, 0, 5)],
        'Ada',
        '2026-09-22',
        []
      );
      peer.applyUpdate(source.encodeState());
      peer.beginUndoCapture();
      source.setCommentRanges(commentId, [range('body', paraId, 6, 12)]);
      peer.applyUpdate(source.encodeStateAsUpdate(peer.encodeStateVector()));
      expect(peer.resolveComment(commentId)).toEqual([{ story: 'body', start: 6, end: 12 }]);
      expect(peer.canUndo()).toBe(false);
      expect(source.undo()).toBe(true);
      peer.applyUpdate(source.encodeStateAsUpdate(peer.encodeStateVector()));
      expect(peer.resolveComment(commentId)).toEqual([{ story: 'body', start: 0, end: 5 }]);
    } finally {
      source.destroy();
      peer.destroy();
    }
  });
});

import { expect } from 'bun:test';

import {
  STORY,
  exchange,
  fingerprint,
  paragraphText,
  replicas,
} from './context';
import type { DocxScenario } from './context';

export const twoEditorsConverge: DocxScenario = {
  name: 'two-editors-converge',
  description:
    'Two sessions type into different paragraphs, then into the same offset of one paragraph, and one inserts a table while the other splits a paragraph; after the updates are delivered both ways every paragraph must read the same on both sides.',
  participants: ['web:a', 'web:b'],
  async run(ctx) {
    const [a, b] = await replicas(ctx, ['web:a', 'web:b']);
    const paragraphs = a.editor.session
      .paragraphs(STORY)
      .filter((paragraph) => /\S/.test(paragraph.text));
    expect(paragraphs.length).toBeGreaterThan(1);
    const first = paragraphs[0];
    const second = paragraphs[paragraphs.length - 1];
    const spans = new Map(
      a.editor.session
        .paragraphSpans(STORY)
        .map((span) => [span.paraId, span.length])
    );

    a.timer.op('insertText:a', () =>
      a.editor.session.insertText(
        {
          story: STORY,
          paraId: first.paraId,
          offset: spans.get(first.paraId)!,
        },
        ' [A]'
      )
    );
    b.timer.op('insertText:b', () =>
      b.editor.session.insertText(
        {
          story: STORY,
          paraId: second.paraId,
          offset: spans.get(second.paraId)!,
        },
        ' [B]'
      )
    );
    expect(exchange([a, b])).toBeGreaterThan(0);
    expect(fingerprint(a.editor.session)).toBe(fingerprint(b.editor.session));
    expect(paragraphText(a.editor.session, second.paraId)).toContain('[B]');
    expect(paragraphText(b.editor.session, first.paraId)).toContain('[A]');

    a.timer.op('insertText:sameSpot', () =>
      a.editor.session.insertText(
        { story: STORY, paraId: first.paraId, offset: 0 },
        'a'
      )
    );
    b.timer.op('insertText:sameSpot', () =>
      b.editor.session.insertText(
        { story: STORY, paraId: first.paraId, offset: 0 },
        'b'
      )
    );
    exchange([a, b], 'applyUpdate:conflict');
    const merged = paragraphText(a.editor.session, first.paraId);
    expect(paragraphText(b.editor.session, first.paraId)).toBe(merged);
    expect(merged.startsWith('ab') || merged.startsWith('ba')).toBe(true);

    const storiesBefore = a.editor.session.storyIds().length;
    const table = a.timer.op('insertTable:a', () =>
      a.editor.session.insertTable(
        { story: STORY, paraId: second.paraId, offset: 0 },
        2,
        3
      )
    );
    const split = b.timer.op('splitParagraph:b', () =>
      b.editor.session.splitParagraph({
        story: STORY,
        paraId: first.paraId,
        offset: 1,
      })
    );
    exchange([a, b], 'applyUpdate:structural');

    expect(a.editor.session.storyIds()).toHaveLength(storiesBefore + 6);
    expect(b.editor.session.storyIds()).toHaveLength(storiesBefore + 6);
    expect(
      table.createdStoryIds.every((storyId) =>
        b.editor.session.storyIds().includes(storyId)
      )
    ).toBe(true);
    expect(
      a.editor.session
        .paragraphs(STORY)
        .some((paragraph) => paragraph.paraId === split.secondParaId)
    ).toBe(true);
    expect(fingerprint(a.editor.session)).toBe(fingerprint(b.editor.session));
    expect([...a.editor.session.encodeStateVector()].join()).toBe(
      [...b.editor.session.encodeStateVector()].join()
    );

    const catchUp = a.timer.op('encodeStateAsUpdate:inSync', () =>
      a.editor.session.encodeStateAsUpdate(b.editor.session.encodeStateVector())
    );
    expect(catchUp.byteLength).toBeLessThan(
      a.editor.session.encodeStateAsUpdate().byteLength
    );
  },
};

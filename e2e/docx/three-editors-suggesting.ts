import { expect } from 'bun:test';

import {
  STORY,
  exchange,
  fingerprint,
  paragraphText,
  replicas,
} from './context';
import type { DocxScenario } from './context';

const DATE = '2026-09-22T00:00:00Z';

export const threeEditorsSuggesting: DocxScenario = {
  name: 'three-editors-suggesting',
  description:
    'Three sessions each suggest an edit as a distinct author, exchange in a full mesh, and every replica must list the same three revisions; one is accepted and one rejected, and the outcome converges again.',
  participants: ['web:a', 'web:b', 'web:c'],
  async run(ctx) {
    const peers = await replicas(ctx, ['web:a', 'web:b', 'web:c']);
    const [a, b, c] = peers;
    const paragraphs = a.editor.session
      .paragraphs(STORY)
      .filter((paragraph) => /\S{4}/.test(paragraph.text));
    expect(paragraphs.length).toBeGreaterThanOrEqual(3);
    const spans = new Map(
      a.editor.session
        .paragraphSpans(STORY)
        .map((span) => [span.paraId, span.length])
    );
    const converged = () => {
      const reference = fingerprint(a.editor.session);
      for (const peer of peers)
        expect(fingerprint(peer.editor.session)).toBe(reference);
      return reference;
    };

    const revisions: string[] = [];
    peers.forEach((peer, index) => {
      const paragraph = paragraphs[index];
      const author = { name: `Author ${index}`, date: DATE };
      const receipt = peer.timer.op('insertText:suggesting', () =>
        peer.editor.session.insertText(
          {
            story: STORY,
            paraId: paragraph.paraId,
            offset: spans.get(paragraph.paraId)!,
          },
          ` [suggested ${index}]`,
          author
        )
      );
      expect(receipt.revisionId).toBeTruthy();
      revisions.push(receipt.revisionId!);
      expect(
        peer.timer
          .op('listRevisions:own', () => peer.editor.session.listRevisions())
          .some((entry) => entry.revisionId === receipt.revisionId)
      ).toBe(true);
    });

    expect(exchange(peers)).toBeGreaterThan(0);
    converged();
    for (const peer of peers) {
      const listed = peer.timer.op('listRevisions:meshed', () =>
        peer.editor.session.listRevisions()
      );
      const ids = listed.map((entry) => entry.revisionId);
      for (const revisionId of revisions) expect(ids).toContain(revisionId);
      expect(
        listed.filter((entry) => entry.kind === 'insertion').length
      ).toBeGreaterThanOrEqual(peers.length);
    }
    expect(paragraphText(a.editor.session, paragraphs[1].paraId)).toContain(
      '[suggested 1]'
    );

    const accepted = a.timer.op('acceptChange', () =>
      a.editor.session.acceptChange({ revisionId: revisions[0] })
    );
    expect(accepted.revisionIds).toContain(revisions[0]);
    const rejected = b.timer.op('rejectChange', () =>
      b.editor.session.rejectChange({ revisionId: revisions[1] })
    );
    expect(rejected.revisionIds).toContain(revisions[1]);
    exchange(peers, 'applyUpdate:afterResolve');

    converged();
    for (const peer of peers) {
      const ids = peer.timer
        .op('listRevisions:afterResolve', () =>
          peer.editor.session.listRevisions()
        )
        .map((entry) => entry.revisionId);
      expect(ids).not.toContain(revisions[0]);
      expect(ids).not.toContain(revisions[1]);
      expect(ids).toContain(revisions[2]);
    }
    expect(paragraphText(c.editor.session, paragraphs[0].paraId)).toContain(
      '[suggested 0]'
    );
    expect(paragraphText(c.editor.session, paragraphs[1].paraId)).not.toContain(
      '[suggested 1]'
    );
  },
};

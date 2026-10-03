import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { createProposalRegistry } from './proposals';
import { createResidentEngineSession } from './residentEngineSession';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURES = [
  resolve(import.meta.dir, '../../../../crates/docx-edit/tests/fixtures/structured-export/principal.docx'),
  resolve(import.meta.dir, '__fixtures__/comment-ranges/structure.docx'),
];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

function comments(session: YrsSession) {
  const result = session.exportStructured({ revisionView: 'markup', stories: ['comments'] });
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

function expectSourceComments(replica: YrsSession, seeded: YrsSession) {
  const expected = comments(seeded);
  const actual = comments(replica);
  for (const result of [expected, actual]) {
    expect(result.content.stories.map((story) => story.comment)).toContainEqual(
      expect.objectContaining({ author: expect.any(String), date: expect.any(String) })
    );
  }
  const comparable = (result: typeof expected) => JSON.stringify({
    ...result.content,
    diagnostics: result.content.diagnostics.filter(({ code }) => code !== 'provenance-unavailable'),
  });
  expect(comparable(actual)).toBe(comparable(expected));
}

function editComment(session: YrsSession, id: string, author: string) {
  const range = session.resolveComment(id)[0]!;
  session.applyRawOps(range.story, [{
    op: 'setComment',
    id,
    ranges: [[range.start, range.end]],
    author,
    date: '2026-10-01T00:00:00Z',
    body: 'Updated comment',
  }]);
}

function expectEditedComment(session: YrsSession, id: string, author: string) {
  const story = comments(session).content.stories.find((story) => story.comment?.id === id)!;
  expect(story.comment).toMatchObject({ author, date: '2026-10-01T00:00:00Z' });
  expect(story.blocks).toMatchObject([{
    paragraph: { inlines: [{ kind: 'text', text: 'Updated comment' }] },
  }]);
}

for (const fixture of FIXTURES) {
  const name = fixture.endsWith('/principal.docx') ? 'principal.docx' : 'structure.docx';

  for (const proposal of [false, true]) {
    test(`a worker replica retains source comments with proposal=${proposal}: ${name}`, async () => {
      const bytes = new Uint8Array(readFileSync(fixture));
      const worker = await createResidentEngineSession();
      const seeded = await createYrsSession({ clientId: 97101 });
      const replica = await createYrsSession({ clientId: 97102 });
      try {
        worker.openDocx(bytes);
        seeded.openDocx(bytes, true);
        if (proposal) {
          const registry = createProposalRegistry(worker.proposalEngine);
          const paragraph = worker.paragraphIdentities().paragraphs.find(
            (paragraph) => paragraph.session?.story === 'body'
          )!.session!;
          expect(registry.propose({
            expectVersion: worker.proposalEngine.version(),
            proposals: [{
              id: 'comment-handover',
              paragraph,
              suggest: { author: 'Host', date: '2026-10-01T00:00:00Z' },
              op: 'insertText',
              at: 'end',
              text: '!',
            }],
          }).ok).toBe(true);
        }
        replica.openDocx(bytes, false);
        replica.loadState(worker.encodeState());
        expectSourceComments(replica, seeded);
      } finally {
        replica.destroy();
        seeded.destroy();
        worker.destroy();
      }
    });
  }

  test(`authored comments in the loaded state survive the replica baseline: ${name}`, async () => {
    const bytes = new Uint8Array(readFileSync(fixture));
    const seeded = await createYrsSession({ clientId: 97105 });
    const replica = await createYrsSession({ clientId: 97106 });
    try {
      seeded.openDocx(bytes, true);
      const id = comments(seeded).content.stories[0]!.comment!.id;
      editComment(seeded, id, 'Source author');
      replica.openDocx(bytes, false);
      replica.loadState(seeded.encodeState());
      expectEditedComment(replica, id, 'Source author');
      const otherComments = (session: YrsSession) => comments(session).content.stories.filter(
        (story) => story.comment?.id !== id
      );
      expect(otherComments(replica)).toEqual(otherComments(seeded));
    } finally {
      replica.destroy();
      seeded.destroy();
    }
  });

  test(`comment updates before the replica baseline survive the first load: ${name}`, async () => {
    const bytes = new Uint8Array(readFileSync(fixture));
    const worker = await createResidentEngineSession();
    const replica = await createYrsSession({ clientId: 97107 });
    const peer = await createYrsSession({ clientId: 97108 });
    try {
      worker.openDocx(bytes);
      peer.openDocx(bytes, false);
      peer.loadState(worker.encodeState());
      const id = comments(peer).content.stories[0]!.comment!.id;
      editComment(peer, id, 'Peer author');
      replica.openDocx(bytes, false);
      replica.applyUpdate(peer.encodeState());
      expectEditedComment(replica, id, 'Peer author');
      replica.loadState(worker.encodeState());
      expectEditedComment(replica, id, 'Peer author');
    } finally {
      peer.destroy();
      replica.destroy();
      worker.destroy();
    }
  });

  test(`comment writes after the replica baseline survive later loads and updates: ${name}`, async () => {
    const bytes = new Uint8Array(readFileSync(fixture));
    const worker = await createResidentEngineSession();
    const replica = await createYrsSession({ clientId: 97103 });
    const peer = await createYrsSession({ clientId: 97104 });
    try {
      worker.openDocx(bytes);
      replica.openDocx(bytes, false);
      replica.loadState(worker.encodeState());
      const id = comments(replica).content.stories[0]!.comment!.id;
      editComment(replica, id, 'Local author');
      expectEditedComment(replica, id, 'Local author');
      replica.loadState(worker.encodeState());
      expectEditedComment(replica, id, 'Local author');

      peer.openDocx(bytes, false);
      peer.loadState(replica.encodeState());
      editComment(peer, id, 'Peer author');
      replica.applyUpdate(peer.encodeState());
      expectEditedComment(replica, id, 'Peer author');
    } finally {
      peer.destroy();
      replica.destroy();
      worker.destroy();
    }
  });
}

import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { TrackedChangesResult } from '@betteroffice/docx/layout/render';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { SidebarRevisionReads } from './sidebarRevisionReads';

beforeAll(async () => {
  const { preloadEditWasm } = await import('@betteroffice/docx/wasm/edit');
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  );
});

const AUTHOR = { name: 'Ada', date: '2026-09-29T00:00:00Z' };

function counting(session: YrsSession): { session: YrsSession; lists: () => number } {
  let lists = 0;
  const counted = Object.create(session) as YrsSession;
  counted.listRevisions = () => {
    lists += 1;
    return session.listRevisions();
  };
  counted.version = () => session.version();
  return { session: counted, lists: () => lists };
}

test('revisions and entries are read once per document version', async () => {
  const live = await createYrsSession({ clientId: 80021 });
  try {
    const [paraId] = live.loadStories([{ storyId: 'body', paragraphs: [{ text: 'alpha' }] }]).body;
    const { session, lists } = counting(live);
    const reads = new SidebarRevisionReads();
    live.insertText({ story: 'body', paraId: paraId!, offset: 5 }, ' beta', AUTHOR);
    const first = reads.tracked(session).tracked;
    expect(reads.tracked(session).tracked).toBe(first);
    expect(reads.revisions(session).revisions).toHaveLength(1);
    expect(lists()).toBe(1);

    live.insertText({ story: 'body', paraId: paraId!, offset: 0 }, 'gamma ', AUTHOR);
    const second = reads.tracked(session).tracked;
    expect(second).not.toBe(first);
    expect(second.entries).toHaveLength(2);
    expect(lists()).toBe(2);
  } finally {
    live.destroy();
  }
});

test('a result delivered at an earlier version is delivered again', async () => {
  const live = await createYrsSession({ clientId: 80022 });
  try {
    const [paraId] = live.loadStories([{ storyId: 'body', paragraphs: [{ text: 'alpha' }] }]).body;
    const reads = new SidebarRevisionReads();
    const received: TrackedChangesResult[] = [];
    const to = (result: TrackedChangesResult): void => {
      received.push(result);
    };
    const empty: TrackedChangesResult = { entries: [], commentToRevision: new Map() };
    reads.deliver(to, empty, live, reads.revisions(live).version);
    reads.deliver(to, empty, live, reads.revisions(live).version);
    expect(received).toHaveLength(1);

    // The editor publishes its own entries after an accept; the document then changes back to
    // no revisions, and the same empty result must replace them.
    const { revisionId } = live.insertText(
      { story: 'body', paraId: paraId!, offset: 5 },
      ' beta',
      AUTHOR
    );
    live.acceptChange({ revisionId: revisionId! });
    reads.deliver(to, empty, live, reads.revisions(live).version);
    expect(received).toEqual([empty, empty]);

    const other: typeof to = () => {};
    reads.deliver(other, empty, live, reads.revisions(live).version);
    reads.deliver(to, empty, live, reads.revisions(live).version);
    expect(received).toHaveLength(3);
  } finally {
    live.destroy();
  }
});

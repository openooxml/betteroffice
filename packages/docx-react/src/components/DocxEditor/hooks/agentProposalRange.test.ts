import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { overlapsTextRevision } from './agentProposalRange';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
)))));

test('detects mixed insertion/deletion ranges and insertion boundaries without blocking unrelated text', async () => {
  const session = await createYrsSession({ clientId: 995 });
  try {
    const { paraId } = session.createStory('body', 'Hello world');
    session.createStory('fn:1', 'Other story');
    session.insertText({ story: 'body', paraId, offset: 5 }, ' new', {
      name: 'First agent', date: '2026-09-13T00:00:00Z',
    });
    const range = (start: number, end: number) => ({
      story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end },
    });
    expect(session.selectionContext(range(0, 9)).inInsertion).toBe(false);
    expect(overlapsTextRevision(session, range(0, 9))).toBe(true);
    expect(overlapsTextRevision(session, range(0, 5))).toBe(false);
    expect(overlapsTextRevision(session, range(5, 5))).toBe(true);
    expect(overlapsTextRevision(session, range(9, 9))).toBe(true);
    expect(overlapsTextRevision(session, range(10, 15))).toBe(false);
    session.rejectChange({ revisionId: session.listRevisions()[0].revisionId });
    session.deleteRange(range(6, 11), { name: 'First agent', date: '2026-09-13T00:00:00Z' });
    expect(session.selectionContext(range(0, 11)).inDeletion).toBe(false);
    expect(overlapsTextRevision(session, range(0, 11))).toBe(true);
    const other = session.paragraphs('fn:1')[0].paraId;
    expect(overlapsTextRevision(session, {
      story: 'fn:1', start: { paraId: other, offset: 0 }, end: { paraId: other, offset: 5 },
    })).toBe(false);
  } finally {
    session.destroy();
  }
});

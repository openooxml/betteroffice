import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { createResidentEngineSession } from './residentEngineSession';
import { createProposalRegistry, type DocxProposalSnapshot } from './proposals';
import { saveYrsDocx } from './saveYrsDocx';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURES = resolve(import.meta.dir, '../../../../crates/docx-edit/tests/fixtures');
const DOCUMENTS = ['page-fragments/pages.docx', 'footnote-anchor.docx', 'structured-export/principal.docx'];

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

function texts(session: YrsSession): Record<string, string[]> {
  return Object.fromEntries(
    session
      .storyIds()
      .sort()
      .map((story) => [story, session.paragraphs(story).map((paragraph) => paragraph.text)])
  );
}

for (const name of DOCUMENTS) {
  test(`a replica of a document the worker opened reads and saves as one opened directly: ${name}`, async () => {
    const bytes = new Uint8Array(readFileSync(resolve(FIXTURES, name)));
    const direct = await createYrsSession({ clientId: 97001 });
    const directHost = direct.openDocx(bytes, true);

    const worker = await createResidentEngineSession();
    const host = decodeDocxHostJson(worker.openDocx(bytes), bytes);
    expect(host.referencedFonts).toEqual(directHost.referencedFonts);
    expect(host.document).toEqual(directHost.document);

    const replica = await createYrsSession({ clientId: 97002 });
    replica.openDocx(bytes, false);
    replica.loadState(worker.encodeState());
    expect(texts(replica)).toEqual(texts(direct));

    const saved = await createYrsSession({ clientId: 97003 });
    saved.openDocx((await saveYrsDocx(replica)).bytes, true);
    const savedDirect = await createYrsSession({ clientId: 97004 });
    savedDirect.openDocx((await saveYrsDocx(direct)).bytes, true);
    expect(texts(saved)).toEqual(texts(savedDirect));

    for (const session of [direct, replica, saved, savedDirect]) session.destroy();
    worker.destroy();
  });
}

test('a worker document mirror survives hydration and hands proposal decisions to the main session', async () => {
  const bytes = new Uint8Array(readFileSync(resolve(FIXTURES, DOCUMENTS[0]!)));
  const worker = await createResidentEngineSession();
  const main = await createYrsSession({ clientId: 97005 });
  const registry = createProposalRegistry(worker.proposalEngine);
  try {
    worker.openDocx(bytes);
    const paragraph = worker.paragraphIdentities().paragraphs.find(
      (paragraph) => paragraph.session?.story === 'body'
    )!.session!;
    const proposed = registry.propose({
      expectVersion: worker.proposalEngine.version(),
      proposals: [{
        id: 'mirrored',
        paragraph,
        suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
        op: 'insertText',
        at: 'end',
        text: '!',
      }],
    });
    expect(proposed.ok).toBe(true);
    registry.setStates({
      expectVersion: worker.proposalEngine.version(),
      expectPreviewVersion: 0,
      changes: [{ id: 'mirrored', state: 'accepted' }],
    });
    const events: DocxProposalSnapshot[] = [];
    main.onProposalChange((snapshot) => {
      expect(main.version()).toBe(snapshot.version);
      events.push(snapshot);
    });
    const mirror = { version: worker.proposalEngine.version(), proposals: registry.exportState() };
    main.mirrorWorkerDocument(mirror);
    expect(main.workerDocumentMirrored()).toBe(true);
    expect(main.version()).toBe(mirror.version);
    expect(main.getProposals()).toEqual(registry.snapshot());
    expect(events).toHaveLength(1);
    main.mirrorWorkerDocument(structuredClone(mirror));
    expect(events).toHaveLength(1);
    main.adoptResidentWorkerLayout!(JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: {},
      renderEnv: {},
    }));
    const snapshot = main.residentWorkerSnapshot({ knownStateVector: new Uint8Array([255]) })!;
    expect(snapshot.workerAuthoritative).toBe(true);
    expect(snapshot.state).toEqual(new Uint8Array(0));
    expect(snapshot.selection).toBeNull();
    expect(snapshot).not.toHaveProperty('mediaSources');
    expect(snapshot).not.toHaveProperty('noteSeparators');

    const state = worker.encodeState();
    main.openDocx(bytes, false);
    expect(main.version()).toBe(mirror.version);
    expect(main.getProposals()).toEqual(registry.snapshot());
    main.loadState(state);
    expect(main.workerDocumentMirrored()).toBe(true);
    expect(main.getProposals()).toEqual(registry.snapshot());
    expect(events).toHaveLength(1);
    main.mirrorWorkerDocument(null);
    expect(main.workerDocumentMirrored()).toBe(false);
    expect(main.getProposals()).toEqual({ ...registry.snapshot(), version: main.version() });
    expect(events).toHaveLength(2);
    const decided = main.setProposalStates({
      expectVersion: main.version(),
      expectPreviewVersion: registry.snapshot().previewVersion,
      changes: [{ id: 'mirrored', state: 'rejected' }],
    });
    expect(decided).toMatchObject({
      ok: true,
      snapshot: { version: main.version(), previewVersion: 2, proposals: [{ state: 'rejected' }] },
    });
    expect(events).toHaveLength(3);
    expect(main.residentWorkerSnapshot()!.workerAuthoritative).toBeUndefined();
  } finally {
    registry.destroy();
    main.destroy();
    worker.destroy();
  }
});

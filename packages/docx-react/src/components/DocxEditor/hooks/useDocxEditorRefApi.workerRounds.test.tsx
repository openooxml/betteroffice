import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { createYrsSession, type DocxProposalInput, type DocxProposalResult, type YrsSession } from '@betteroffice/docx/yrs';
import { createStyleResolver } from '@betteroffice/docx/styles';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import { applyEditBatch } from '../editorBatches';
import { installEditorWorkerProposalActivation, registeredWorkerProposalAuthority, handedOverRequest, registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import { deferWorkerOpenReplica, requestWorkerOpenReplica } from '../internals/workerOpenReplica';
import { useRustDisplayList } from './useDisplayList';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';
import { useRevisionPreview } from './useRevisionPreview';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;
const sessions: YrsSession[] = [];
const workers: InProcessResidentWorker[] = [];
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
let nextClientId = 98500;
const LAYOUT = JSON.stringify({
  bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontFamily: 'Calibri', fontSize: 11 } }, renderEnv: {},
});
const SUGGEST = { author: 'Host', date: '2026-10-05T00:00:00Z' };

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});
afterEach(() => {
  cleanup();
  mock.restore();
  for (const worker of workers.splice(0)) worker.terminate();
  for (const session of sessions.splice(0)) session.destroy();
  globalThis.Worker = originalWorker;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((yes) => { resolve = yes; });
  return { promise, resolve };
}

function documentBytes() {
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Alpha</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function text(session: Pick<YrsSession, 'readParagraphs'>) {
  const read = session.readParagraphs({ view: 'accepted' });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map(({ text }) => text).join('\n');
}

function snapshot(result: DocxProposalResult) {
  expect(result.ok).toBe(true);
  if (!result.ok) throw new Error(result.failure.message);
  return result.snapshot;
}

async function editor(holdHydration = false) {
  globalThis.Worker = class {
    constructor() { const worker = startWorker(nextClientId++); workers.push(worker); return worker; }
  } as unknown as typeof Worker;
  const bytes = documentBytes();
  const peer = await createYrsSession({ clientId: nextClientId++ });
  sessions.push(peer);
  const loaded = deferred<void>();
  const release = deferred<void>();
  const events: string[] = [];
  const refresh = mock(() => true);
  const peerProjection = mock((_stories: readonly string[]) => {});
  const renderer = renderHook(() => useRustDisplayList(
    null, undefined, undefined, undefined, peer, undefined, undefined, undefined, true
  ));
  const opened = await renderer.result.current.openInWorker(peer, bytes);
  if (!opened) throw new Error('No worker document');
  const worker = workers.at(-1)!;
  const { document } = peer.openDocx(bytes, false);
  const replica = deferWorkerOpenReplica(peer, async () => {
    const state = await opened.encodeVersionedState!();
    events.push('snapshot');
    if (holdHydration) { loaded.resolve(); await release.promise; }
    return [
      () => { peer.loadState(state.state); events.push('load'); loaded.resolve(); },
    ];
  }, () => { throw new Error('unexpected fallback'); }, () => {
    events.push('ready');
    peer.beginUndoCapture();
    opened.replicaReady();
  }, {
    current: () => true, cancel: () => {},
  });
  await renderer.result.current.layoutInWorker(peer, LAYOUT);
  let authority: ReturnType<typeof registerWorkerProposalAuthority> | null = null;
  installEditorWorkerProposalActivation(peer, () => authority ??= registerWorkerProposalAuthority(peer, opened, {
    editorPeer: true, current: () => true, laidOut: async () => {}, adopted: () => {},
    relayout: () => {}, contentChanged: () => {}, projectionChanged: peerProjection,
    peerUpdated: () => { if (!replica.pending) refresh(); },
  }));
  const editor = {
    getYrsSession: () => peer, getDocument: () => document,
    flushPendingInput: async () => {}, syncYrsInputState: refresh,
    isWorkerViewer: () => false, getLayout: () => null,
  } as unknown as PagedEditorRef;
  const pagedEditorRef = { current: editor };
  const modeRef = { current: 'editing' as 'editing' | 'viewing' };
  const allowHostProposalsRef = { current: true };
  const hook = renderHook(() => {
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      experimentalWorkerOpen: true, ref, document, documentFromYrs: () => document,
      historyStateRef: { current: document }, pagedEditorRef,
      handleSave: async () => null, zoom: 1, setZoom: () => {},
      scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
      loadParsedDocument: () => {}, loadBuffer: async () => {}, comments: [],
      setComments: () => {}, setShowCommentsSidebar: () => {},
      contentChangeSubscribersRef: { current: new Set() }, selectionChangeSubscribersRef: { current: new Set() },
      getCachedStyleResolver: createStyleResolver, commentIdAllocator: createCommentIdAllocator(),
      commands: UNAVAILABLE_DOCX_COMMANDS, modeRef, allowHostProposalsRef,
      hostSearch: {
        search: async () => ({ query: '', options: { caseSensitive: false }, total: 0, current: -1 }),
        searchNext: () => null, searchPrevious: () => null, searchGoTo: () => null, clearSearch: () => {},
        getSearchState: () => null, onSearchChange: () => () => {},
      },
    });
    return { api: ref, preview: useRevisionPreview(peer) };
  });
  const ready = requestWorkerOpenReplica(peer)!;
  if (holdHydration) await loaded.promise;
  else {
    await act(async () => { await ready; });
    snapshot(await hook.result.current.api.current!.proposeChanges({ expectVersion: peer.version(), proposals: [] }));
  }
  const proposal = (id: string, content = 'Worker ', at: 'start' | 'end' = 'start'): DocxProposalInput => ({
    id, paragraph: {
      kind: 'persisted', story: { kind: 'body', partUri: '/word/document.xml' }, paraId: '00000001',
    }, suggest: SUGGEST, op: 'insertText', at, text: content,
  });
  const type = (content = 'Typed ') => {
    peer.insertText({ story: 'body', paraId: '00000001', offset: 0 }, content);
  };
  return {
    peer, worker, get authority() {
      const current = registeredWorkerProposalAuthority(peer);
      if (!current) throw new Error('No editor round has activated');
      return current;
    }, replica, ready, release: () => release.resolve(), events, refresh, peerProjection,
    api: hook.result.current.api.current!, hook, renderer, pagedEditorRef, proposal, type, modeRef, allowHostProposalsRef,
    workerText: () => text(worker.sessions[0]!.proposalEngine),
  };
}

test('ready editor ref rounds execute in the worker and integrate without whole-state decoding', async () => {
  const h = await editor();
  const propose = spyOn(h.peer, 'proposeChanges');
  const encode = spyOn(h.peer, 'encodeState');
  const load = spyOn(h.peer, 'loadState');
  const host = spyOn(h.peer, 'applyHostUpdate');
  const local = spyOn(h.peer, 'applyLocalUpdate');
  const result = snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('round')] }));
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(3);
  expect(propose).not.toHaveBeenCalled();
  expect(host).toHaveBeenCalledTimes(1);
  expect(local).not.toHaveBeenCalled();
  expect(encode).not.toHaveBeenCalled();
  expect(load).not.toHaveBeenCalled();
  expect(result.version).toBe(h.peer.version());
  expect(text(h.peer)).toBe('Worker Alpha');
  expect(text(h.peer)).toBe(h.workerText());
  expect(h.refresh).toHaveBeenCalledTimes(1);
  expect(h.peerProjection).toHaveBeenLastCalledWith(['body']);
});

test('integrating a round diff emits no applyUpdate echo', async () => {
  const h = await editor();
  const before = h.worker.requests.filter((type) => type === 'applyUpdate').length;
  snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('echo')] }));
  await h.authority.residentOperation(async () => {});
  expect(h.worker.requests.filter((type) => type === 'applyUpdate')).toHaveLength(before);
  expect(text(h.peer)).toBe(h.workerText());
});

test('peer undo and redo retain worker proposal content while reverting only typing', async () => {
  const h = await editor();
  h.type();
  snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('undo')] }));
  expect(text(h.peer)).toContain('Typed ');
  expect(h.peer.undo()).toBe(true);
  expect(text(h.peer)).toBe('Worker Alpha');
  expect(h.peer.redo()).toBe(true);
  expect(text(h.peer)).toContain('Typed ');
  expect(text(h.peer)).toContain('Worker ');
  expect((await h.api.getProposals()).proposals.map(({ id }) => id)).toEqual(['undo']);
});

test('typing before a queued round posts refuses its stale token', async () => {
  const h = await editor();
  const blocker = deferred<void>();
  const started = deferred<void>();
  const held = h.authority.residentOperation(async () => { started.resolve(); await blocker.promise; });
  await started.promise;
  const token = h.peer.version();
  const pending = h.api.proposeChanges({ expectVersion: token, proposals: [h.proposal('stale')] });
  h.type();
  const before = h.worker.requests.filter((type) => type === 'proposal').length;
  blocker.resolve();
  await held;
  expect(await pending).toMatchObject({ ok: false, version: h.peer.version(), failure: { code: 'stale-version' } });
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before);
  expect((await h.api.getProposals()).proposals).toEqual([]);
});

test('typing after posting merges with a successful round and invalidates coverage', async () => {
  const h = await editor();
  h.worker.hold();
  const posted = deferred<void>();
  const send = h.worker.postMessage.bind(h.worker);
  spyOn(h.worker, 'postMessage').mockImplementation((message, transfer) => {
    send(message, transfer);
    if ('type' in message && message.type === 'proposal') posted.resolve();
  });
  const pending = h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('merge')] });
  await posted.promise;
  h.type();
  h.worker.release();
  const result = snapshot(await pending);
  expect(result.version).toBe(h.peer.version());
  expect(text(h.peer)).toContain('Typed ');
  expect(text(h.peer)).toContain('Worker ');
  expect(h.authority.workerCoversPeer(h.peer.version())).toBe(false);
  await h.api.getProposals();
  expect(text(h.peer)).toBe(h.workerText());
});

test('the first ref round waits for base hydration and then runs in the worker without fallback', async () => {
  const h = await editor(true);
  expect(h.replica.pending).toBe(true);
  const before = h.worker.requests.filter((type) => type === 'proposal').length;
  const peerPropose = spyOn(h.peer, 'proposeChanges');
  const round = h.api.proposeChanges({ expectVersion: h.worker.sessions[0]!.proposalEngine.version(), proposals: [h.proposal('during')] });
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before);
  expect(h.events).toEqual(['snapshot']);
  expect(h.replica.pending).toBe(true);
  expect(registeredWorkerProposalAuthority(h.peer)).toBeNull();
  h.release();
  await act(async () => { await h.ready; snapshot(await round); });
  expect(peerPropose).not.toHaveBeenCalled();
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before + 2);
  expect(h.events).toEqual(['snapshot', 'load', 'ready']);
  expect(text(h.peer)).toBe(h.workerText());
  expect(h.authority.workerCoversPeer(h.peer.version())).toBe(true);
  expect(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('after', 'After ')] }))
    .toMatchObject({ ok: false, failure: { code: 'tracked-revision-conflict' } });
  snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('after', 'After ', 'end')] }));
  expect(text(h.peer)).toContain('After ');
  expect(text(h.peer)).toBe(h.workerText());
});

test('same proposal ids stay independent and both decisions enter the preview by revision id', async () => {
  const h = await editor();
  const local = snapshot(h.peer.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('same', 'Local ')] }));
  expect(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('same')] }))
    .toMatchObject({ ok: false, failure: { code: 'tracked-revision-conflict' } });
  const remote = snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('same', 'Worker ', 'end')] }));
  expect(h.peer.getProposals().proposals).toEqual(local.proposals);
  expect((await h.api.getProposals()).proposals).toEqual(remote.proposals);
  const localIds = local.proposals.flatMap(({ revisionIds }) => revisionIds);
  const workerIds = remote.proposals.flatMap(({ revisionIds }) => revisionIds);
  expect(localIds.length).toBeGreaterThan(0);
  expect(workerIds.length).toBeGreaterThan(0);
  expect(localIds.some((id) => workerIds.includes(id))).toBe(false);
  await act(async () => {
    snapshot(h.peer.setProposalStates({ expectVersion: h.peer.version(), expectPreviewVersion: local.previewVersion, changes: [{ id: 'same', state: 'rejected' }] }));
    snapshot(await h.api.setProposalStates({ expectVersion: h.peer.version(), expectPreviewVersion: remote.previewVersion, changes: [{ id: 'same', state: 'accepted' }] }));
  });
  const preview = h.hook.result.current.preview.revisionPreview!;
  expect(Object.keys(preview).sort()).toEqual([...localIds, ...workerIds].sort());
  for (const id of localIds) expect(preview[id]).toBe('rejected');
  for (const id of workerIds) expect(preview[id]).toBe('accepted');
  expect(h.authority.revisionPreview()).toEqual(preview);
  const settled = snapshot(await h.api.withdrawProposals({ expectVersion: h.peer.version(), ids: ['same'] }));
  expect(settled.proposals).toEqual([]);
  expect(h.peer.getProposals().proposals.map(({ id }) => id)).toEqual(['same']);
});

test('same proposal ids retain separate records and decisions after permanent worker loss', async () => {
  const h = await editor();
  let local!: ReturnType<typeof snapshot>;
  let remote!: ReturnType<typeof snapshot>;
  await act(async () => {
    local = snapshot(h.peer.proposeChanges({ expectVersion: h.peer.version(), proposals: [{
      ...h.proposal('same', 'Local '), suggest: { ...SUGGEST, author: 'Peer author' },
    }] }));
    remote = snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [{
      ...h.proposal('same', 'Worker ', 'end'), suggest: { ...SUGGEST, author: 'Worker author' },
    }] }));
    local = snapshot(h.peer.setProposalStates({ expectVersion: h.peer.version(),
      expectPreviewVersion: local.previewVersion, changes: [{ id: 'same', state: 'rejected' }] }));
    remote = snapshot(await h.api.setProposalStates({ expectVersion: h.peer.version(),
      expectPreviewVersion: remote.previewVersion, changes: [{ id: 'same', state: 'accepted' }] }));
  });
  const localRecords = structuredClone(local.proposals);
  const workerRecords = structuredClone(remote.proposals);
  const localIds = localRecords.flatMap(({ revisionIds }) => revisionIds);
  const workerIds = workerRecords.flatMap(({ revisionIds }) => revisionIds);
  expect(localIds.length).toBeGreaterThan(0);
  expect(workerIds.length).toBeGreaterThan(0);
  const authors = () => h.peer.listRevisions().filter(({ revisionId }) =>
    [...localIds, ...workerIds].includes(revisionId)).map(({ revisionId, author, date }) => ({ revisionId, author, date }));
  const originalAuthors = authors();
  expect(originalAuthors.some(({ author }) => author === 'Peer author')).toBe(true);
  expect(originalAuthors.some(({ author }) => author === 'Worker author')).toBe(true);
  const previewVersion = h.authority.previewVersion();
  const preview = structuredClone(h.authority.revisionPreview());
  spyOn(console, 'error').mockImplementation(() => {});
  const requests = h.worker.requests.filter((type) => type === 'proposal').length;
  await act(async () => {
    h.worker.onerror?.({ message: 'same-id worker crashed' } as ErrorEvent);
    expect(await h.renderer.result.current.layoutInWorker(h.peer, LAYOUT)).toBeNull();
  });
  expect(h.authority.retirementReason()).toBe('source-fallback');
  expect(h.peer.getProposals().proposals).toEqual(localRecords);
  expect((await h.api.getProposals()).proposals).toEqual(workerRecords);
  expect(authors()).toEqual(originalAuthors);
  expect(h.authority.previewVersion()).toBeGreaterThanOrEqual(previewVersion);
  expect(h.hook.result.current.preview.revisionPreview).toEqual(preview);
  for (const id of localIds) expect(preview![id]).toBe('rejected');
  for (const id of workerIds) expect(preview![id]).toBe('accepted');
  await act(async () => {
    remote = snapshot(await h.api.setProposalStates({ expectVersion: h.peer.version(),
      expectPreviewVersion: remote.previewVersion, changes: [{ id: 'same', state: 'rejected' }] }));
  });
  expect(h.peer.getProposals().proposals).toEqual(localRecords);
  expect(remote.proposals).toEqual(workerRecords.map((record) => ({ ...record, state: 'rejected' })));
  await act(async () => {
    local = snapshot(h.peer.setProposalStates({ expectVersion: h.peer.version(),
      expectPreviewVersion: local.previewVersion, changes: [{ id: 'same', state: 'accepted' }] }));
  });
  expect((await h.api.getProposals()).proposals).toEqual(remote.proposals);
  expect(local.proposals).toEqual(localRecords.map((record) => ({ ...record, state: 'accepted' })));
  await act(async () => {
    snapshot(await h.api.withdrawProposals({ expectVersion: h.peer.version(), ids: ['same'] }));
  });
  expect((await h.api.getProposals()).proposals).toEqual([]);
  expect(h.peer.getProposals().proposals).toEqual(local.proposals);
  await act(async () => {
    snapshot(h.peer.withdrawProposals({ expectVersion: h.peer.version(), ids: ['same'] }));
  });
  expect(h.peer.getProposals().proposals).toEqual([]);
  expect((await h.api.getProposals()).proposals).toEqual([]);
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(requests);
});

test('plugin peer batches map a worker token through every correspondence and reject it after an edit', async () => {
  const h = await editor();
  snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('plugin')] }));
  const token = h.authority.geometry()!.version;
  const workerVersion = h.worker.sessions[0]!.proposalEngine.version();
  expect(token).toBe(h.peer.version());
  expect(handedOverRequest(h.peer, { expectVersion: workerVersion }).expectVersion).toBe(h.peer.version());
  const step = { op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: '00000001' }, at: 'end', text: 'Plugin' } as const;
  const applied = await applyEditBatch(h.pagedEditorRef, () => 'editing', { expectVersion: workerVersion, steps: [step] });
  expect(applied).toMatchObject({ result: { ok: true, applied: true } });
  expect(text(h.peer)).toContain('Plugin');
  expect(handedOverRequest(h.peer, { expectVersion: workerVersion }).expectVersion).toBe(workerVersion);
  const stale = await applyEditBatch(h.pagedEditorRef, () => 'editing', { expectVersion: workerVersion, steps: [step] });
  expect(stale).toMatchObject({ result: { ok: false, failure: { code: 'stale-version' } } });
});


test.each(['propose', 'setStates', 'withdraw', 'snapshot'] as const)(
  'ready %s rounds reserve their position while composition flushes', async (kind) => {
    const h = await editor();
    const started = deferred<void>();
    const composition = deferred<void>();
    spyOn(h.pagedEditorRef.current, 'flushPendingInput').mockImplementationOnce(async () => {
      started.resolve();
      await composition.promise;
      h.type('Composed ');
    });
    const token = h.peer.version();
    const before = h.worker.requests.filter((type) => type === 'proposal').length;
    const pending = kind === 'propose'
      ? h.api.proposeChanges({ expectVersion: token, proposals: [h.proposal('composition')] })
      : kind === 'setStates'
        ? h.api.setProposalStates({ expectVersion: token, expectPreviewVersion: 0, changes: [] })
        : kind === 'withdraw'
          ? h.api.withdrawProposals({ expectVersion: token, ids: ['composition'] })
          : h.api.getProposals();
    const save = mock(async () => text(h.peer));
    const saving = h.authority.save(save);
    await started.promise;
    expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before);
    expect(save).not.toHaveBeenCalled();
    composition.resolve();
    const result = await pending;
    if (kind === 'snapshot') expect(result).toMatchObject({ version: h.peer.version(), proposals: [] });
    else expect(result).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(await saving).toBe('Composed Alpha');
  }
);

test('accepted queued input is flushed before ready round version validation', async () => {
  const h = await editor();
  const token = h.peer.version();
  const flush = spyOn(h.pagedEditorRef.current, 'flushPendingInput').mockImplementationOnce(async () => {
    h.type('Accepted ');
  });
  expect(await h.api.proposeChanges({ expectVersion: token, proposals: [h.proposal('accepted')] }))
    .toMatchObject({ ok: false, version: h.peer.version(), failure: { code: 'stale-version' } });
  expect(flush).toHaveBeenCalledTimes(1);
  expect(text(h.peer)).toBe('Accepted Alpha');
});

test('a queued ready mutation rechecks permission when its queue position runs', async () => {
  const h = await editor();
  const started = deferred<void>();
  const release = deferred<void>();
  const held = h.authority.residentOperation(async () => { started.resolve(); await release.promise; });
  await started.promise;
  const pending = h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('denied')] });
  const before = h.worker.requests.filter((type) => type === 'proposal').length;
  h.modeRef.current = 'viewing';
  h.allowHostProposalsRef.current = false;
  release.resolve();
  await held;
  expect(await pending).toMatchObject({ ok: false, failure: { code: 'read-only' } });
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before);
  expect(text(h.peer)).toBe('Alpha');
});

test('ready round admission rejects a replaced document after input flush', async () => {
  const h = await editor();
  spyOn(h.pagedEditorRef.current, 'flushPendingInput').mockImplementationOnce(async () => {
    h.pagedEditorRef.current = { getYrsSession: () => null } as unknown as PagedEditorRef;
  });
  const before = h.worker.requests.filter((type) => type === 'proposal').length;
  await expect(h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('replaced')] }))
    .rejects.toThrow('document changed');
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before);
});

test('an identical ref retry accepts its original token while a fresh edit refuses it', async () => {
  const h = await editor();
  const request = { expectVersion: h.peer.version(), proposals: [h.proposal('retry')] };
  const first = snapshot(await h.api.proposeChanges(request));
  expect(snapshot(await h.api.proposeChanges(request))).toEqual(first);
  expect(text(h.peer)).toBe('Worker Alpha');
  expect(await h.api.proposeChanges({ ...request, proposals: [h.proposal('fresh', 'Fresh ', 'end')] }))
    .toMatchObject({ ok: false, failure: { code: 'stale-version' } });
});

test('an empty ref round accepts an old token after a peer edit', async () => {
  const h = await editor();
  const version = h.peer.version();
  h.type();
  const before = h.worker.requests.filter((type) => type === 'proposal').length;
  expect(await h.api.proposeChanges({ expectVersion: version, proposals: [] }))
    .toMatchObject({ ok: true, snapshot: { version: h.peer.version(), proposals: [] } });
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before + 1);
  expect(text(h.peer)).toBe('Typed Alpha');
});

test('a round reply arriving after typing invalidates geometry until a caught-up snapshot', async () => {
  const h = await editor();
  const posted = deferred<void>();
  const send = h.worker.postMessage.bind(h.worker);
  spyOn(h.worker, 'postMessage').mockImplementation((message, transfer) => {
    if (message.type === 'proposal' && message.operation.kind === 'propose') {
      h.worker.hold();
      posted.resolve();
    }
    send(message, transfer);
  });
  const round = h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('geometry')] });
  await posted.promise;
  h.type();
  h.worker.release();
  snapshot(await round);
  expect(h.authority.geometry()).toBeNull();
  await h.api.getProposals();
  expect(h.authority.geometry()?.version).toBe(h.peer.version());
  expect(text(h.peer)).toContain('Typed ');
  expect(text(h.peer)).toContain('Worker ');
});

test('a worker crash after a completed round hands over once and keeps edits and proposals', async () => {
  const h = await editor();
  h.type();
  let existing!: ReturnType<typeof snapshot>;
  await act(async () => {
    existing = snapshot(h.peer.proposeChanges({
      expectVersion: h.peer.version(), proposals: [h.proposal('peer-existing', 'Existing ', 'end')],
    }));
  });
  const expected = snapshot(await h.api.proposeChanges({
    expectVersion: h.peer.version(), proposals: [h.proposal('retained')],
  }));
  const before = text(h.peer);
  const report = spyOn(console, 'error').mockImplementation(() => {});
  const local = spyOn(h.peer, 'proposeChanges');
  const requests = h.worker.requests.filter((type) => type === 'proposal').length;
  await act(async () => {
    h.worker.onerror?.({ message: 'completed round worker crashed' } as ErrorEvent);
    expect(await h.renderer.result.current.layoutInWorker(h.peer, LAYOUT)).toBeNull();
  });
  expect(h.authority.retirementReason()).toBe('source-fallback');
  expect(report).toHaveBeenCalledTimes(1);
  expect(workers).toHaveLength(1);
  expect(text(h.peer)).toBe(before);
  expect(await h.api.getProposals()).toMatchObject({ proposals: expected.proposals });
  expect(h.peer.getProposals().proposals).toEqual(existing.proposals);
  await act(async () => {
    snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [{
      ...h.proposal('local'), op: 'insertText', at: { offset: text(h.peer).indexOf('Alpha') + 2 }, text: 'Peer ',
    }] }));
  });
  expect(local).not.toHaveBeenCalled();
  expect(text(h.peer)).toContain('Typed ');
  expect(text(h.peer)).toContain('Worker ');
  expect(text(h.peer)).toContain('Peer ');
  expect(text(h.peer)).toContain('Existing ');
  await act(async () => {
    expect(await h.renderer.result.current.layoutInWorker(h.peer, LAYOUT)).toBeNull();
  });
  expect(h.authority.retire('source-fallback')).toBe(false);
  expect(report).toHaveBeenCalledTimes(1);
  expect(workers).toHaveLength(1);
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(requests);
  expect((await h.api.getProposals()).proposals.map(({ id }) => id)).toEqual(['retained', 'local']);
});

test('a permanent hydrated worker layout drop retires proposal rounds to the peer', async () => {
  const h = await editor();
  const error = new Error('non-terminal layout failure');
  const send = h.worker.postMessage.bind(h.worker);
  spyOn(h.worker, 'postMessage').mockImplementation((message, transfer) => {
    if (message.type === 'sync') queueMicrotask(() => h.worker.onmessage?.({
      data: { id: message.id, ok: false, error: error.message },
    } as MessageEvent));
    else send(message, transfer);
  });
  spyOn(console, 'error').mockImplementation(() => {});
  await act(async () => { expect(await h.renderer.result.current.layoutInWorker(h.peer, LAYOUT)).toBeNull(); });
  expect(h.authority.retirementReason()).toBe('source-fallback');
  expect(h.authority.workerCoversPeer(h.peer.version())).toBe(false);
  const before = h.worker.requests.filter((type) => type === 'proposal').length;
  const peer = spyOn(h.peer, 'proposeChanges');
  snapshot(await h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('fallback')] }));
  expect(peer).not.toHaveBeenCalled();
  expect((await h.api.getProposals()).proposals.map(({ id }) => id)).toEqual(['fallback']);
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before);
  expect(text(h.peer)).toBe('Worker Alpha');
});


test('ready round admission rechecks permission after an active input flush', async () => {
  const h = await editor();
  const flushing = deferred<void>();
  const release = deferred<void>();
  spyOn(h.pagedEditorRef.current, 'flushPendingInput').mockImplementationOnce(async () => {
    flushing.resolve();
    await release.promise;
  });
  const before = h.worker.requests.filter((type) => type === 'proposal').length;
  const round = h.api.proposeChanges({ expectVersion: h.peer.version(), proposals: [h.proposal('mode-during-flush')] });
  await flushing.promise;
  h.modeRef.current = 'viewing';
  h.allowHostProposalsRef.current = false;
  release.resolve();
  expect(await round).toMatchObject({ ok: false, failure: { code: 'read-only' } });
  expect(h.worker.requests.filter((type) => type === 'proposal')).toHaveLength(before);
});

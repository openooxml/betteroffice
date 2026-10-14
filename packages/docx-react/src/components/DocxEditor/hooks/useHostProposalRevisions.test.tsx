import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import { proposalSetIdentity, type DocxProposalSnapshot, type ResidentProposalReply, type YrsSession } from '@betteroffice/docx/yrs';
import { createProposalRegistry, type DocxProposalRegistry, type DocxProposalSession } from '@betteroffice/docx/yrs/proposals';
import { yrsIdToNumericId } from '@betteroffice/docx/layout/render';
import { useHostProposalRevisions } from './useHostProposalRevisions';
import { registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import type { WorkerOpenedDocument } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function proposalSession() {
  let snapshot: DocxProposalSnapshot = { version: 'v1', previewVersion: 0, proposals: [] };
  const listeners = new Set<(snapshot: DocxProposalSnapshot) => void>();
  const session = {
    getProposals: () => snapshot,
    onProposalChange(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  } satisfies Pick<YrsSession, 'getProposals' | 'onProposalChange'>;
  return {
    session: session as YrsSession,
    listeners,
    publish(revisionIds: readonly string[], previewVersion = 0) {
      snapshot = {
        version: 'v1',
        previewVersion,
        proposals:
          revisionIds.length === 0
            ? []
            : [
                {
                  id: 'p1',
                  state: previewVersion === 0 ? 'proposed' : 'accepted',
                  paragraph: { kind: 'session', sessionId: 's1', story: 'body', paraId: 'p1' },
                  revisionIds,
                  changed: true,
                },
              ],
      };
      for (const listener of listeners) listener(snapshot);
    },
  };
}

test('proposal keys update on notifications, stay stable for decisions and follow the current session', () => {
  const first = proposalSession();
  const second = proposalSession();
  const view = renderHook(
    ({ session }: { session: YrsSession | null }) => useHostProposalRevisions(session),
    { initialProps: { session: first.session as YrsSession | null } }
  );
  expect(view.result.current.size).toBe(0);
  expect(first.listeners.size).toBe(1);

  act(() => first.publish(['host-a', '7', '7']));
  expect(view.result.current).toEqual(
    new Set([`revision-${yrsIdToNumericId('host-a')}`, 'revision-7'])
  );
  const keys = view.result.current;
  act(() => first.publish(['7', 'host-a'], 1));
  expect(view.result.current).toBe(keys);

  second.publish(['host-b']);
  view.rerender({ session: second.session });
  expect(view.result.current).toEqual(new Set([`revision-${yrsIdToNumericId('host-b')}`]));
  expect(first.listeners.size).toBe(0);
  expect(second.listeners.size).toBe(1);
  const secondKeys = view.result.current;
  act(() => first.publish(['old-session']));
  expect(view.result.current).toBe(secondKeys);

  act(() => second.publish([]));
  expect(view.result.current.size).toBe(0);
  const empty = view.result.current;
  act(() => second.publish([], 1));
  expect(view.result.current).toBe(empty);
  view.rerender({ session: null });
  expect(view.result.current).toBe(empty);
  expect(second.listeners.size).toBe(0);
});

test('hidden host revisions follow peer and worker registries through activation and retirement', async () => {
  const peer = proposalSession();
  peer.publish(['7']);
  const retiredRegistries: DocxProposalRegistry[] = [];
  const registrySession: DocxProposalSession = {
    version: () => 'v1',
    resolveParagraphAnchor: () => { throw new Error('unexpected anchor read'); },
    findText: () => { throw new Error('unexpected text read'); },
    readParagraphs: () => { throw new Error('unexpected paragraph read'); },
    applyEdits: () => { throw new Error('unexpected edit'); },
    listRevisions: () => [],
    settleRevisions: () => {},
  };
  Object.assign(peer.session, {
    version: registrySession.version,
    encodeStateVector: () => new Uint8Array(),
    storiesChangedSince: () => ({ revision: 0, stories: [] }),
    createWorkerProposalRegistry: (state: Parameters<YrsSession['createWorkerProposalRegistry']>[0]) => {
      const registry = createProposalRegistry(registrySession);
      registry.mirror({ version: 'v1', proposals: state });
      registry.mirror(null);
      retiredRegistries.push(registry);
      return registry;
    },
  } satisfies Pick<YrsSession, 'version' | 'encodeStateVector' | 'storiesChangedSince' | 'createWorkerProposalRegistry'>);
  const allKeys = ['revision-7', 'revision-11', 'revision-13', 'revision-17'];
  const view = renderHook(() => {
    const keys = useHostProposalRevisions(peer.session);
    const allowHostProposals = true;
    const showHostProposalsInSidebar = false;
    const sidebar = allowHostProposals && !showHostProposalsInSidebar
      ? allKeys.filter((key) => !keys.has(key))
      : allKeys;
    return { keys, sidebar };
  });
  expect(view.result.current.keys).toEqual(new Set(['revision-7']));
  expect(view.result.current.sidebar).toEqual(['revision-11', 'revision-13', 'revision-17']);
  const workerReply = (revisionIds: readonly string[]): ResidentProposalReply => {
    const snapshot: DocxProposalSnapshot = {
      version: 'v1', previewVersion: 0,
      proposals: revisionIds.length === 0 ? [] : [{
        id: 'p1', state: 'proposed', changed: true,
        paragraph: { kind: 'session', sessionId: 's1', story: 'body', paraId: 'p1' },
        revisionIds,
      }],
    };
    return {
      mirror: { version: 'v1', proposals: {
        previewVersion: 0,
        entries: snapshot.proposals.map((record) => ({
          record, key: 'worker', suggest: { author: 'Host', date: '2026-10-06T00:00:00Z' },
        })),
      } },
      result: { ok: true, snapshot },
      changedStories: [], updates: [], stateVector: new Uint8Array(), peerDiff: new Uint8Array(),
      geometry: { version: 'v1', previewVersion: 0, proposals: proposalSetIdentity(snapshot), targets: {}, hidden: [] },
    };
  };
  const worker = {
    proposal: async (op) => workerReply(op.kind === 'snapshot' ? [] : ['11', '13']),
    documentRead: async () => { throw new Error('unexpected worker read'); },
    handOver: async () => { throw new Error('unexpected handover'); },
    integrateProposalUpdate: () => [],
  } satisfies Pick<WorkerOpenedDocument, 'proposal' | 'documentRead' | 'handOver' | 'integrateProposalUpdate'>;
  let authority!: ReturnType<typeof registerWorkerProposalAuthority>;
  const readProposals = spyOn(peer.session, 'getProposals');
  act(() => {
    authority = registerWorkerProposalAuthority(peer.session, worker, {
      editorPeer: true, current: () => true, laidOut: async () => {},
      relayout: () => {}, adopted: () => {}, contentChanged: () => {},
    });
  });
  expect(readProposals).toHaveBeenCalledTimes(1);
  readProposals.mockRestore();
  await act(async () => {
    await authority.initialize();
    expect(await authority.propose({ expectVersion: 'v1', proposals: [{
      id: 'p1', paragraph: { kind: 'session', sessionId: 's1', story: 'body', paraId: 'p1' },
      suggest: { author: 'Host', date: '2026-10-06T00:00:00Z' },
      op: 'insertText', at: 'start', text: 'Proposed ',
    }] }, async () => { throw new Error('unexpected peer proposal'); })).toMatchObject({ ok: true });
  });
  expect(peer.session.getProposals().proposals[0]!.revisionIds).toEqual(['7']);
  expect(view.result.current.keys).toEqual(new Set(['revision-7', 'revision-11', 'revision-13']));
  expect(view.result.current.sidebar).toEqual(['revision-17']);
  act(() => { expect(authority.retire('source-fallback')).toBe(true); });
  expect(view.result.current.keys).toEqual(new Set(['revision-7', 'revision-11', 'revision-13']));
  expect(view.result.current.sidebar).toEqual(['revision-17']);
  act(() => peer.publish(['7', '17']));
  expect(view.result.current.sidebar).toEqual([]);
  act(() => retiredRegistries[0]!.reset());
  expect(view.result.current.keys).toEqual(new Set(['revision-7', 'revision-17']));
  expect(view.result.current.sidebar).toEqual(['revision-11', 'revision-13']);
  view.unmount();
  expect(peer.listeners.size).toBe(0);
});

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { usePeerlessCollaboration, createPeerlessCollaborationReplica } from './peerlessCollaboration';
import { PeerlessRemotePresence } from './PeerlessRemotePresence';
import { createPeerlessViewStore } from './viewStore';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { CollaborationPresence, CollaborationReplica } from '@betteroffice/docx/collaboration';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { DocxEditorCollaborationOptions } from '../types';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });
const source = new Uint8Array([1]);

function setup() {
  const listeners = new Set<(event: { updates: Uint8Array[]; inferences: [] }) => void>();
  const client = { onCollabUpdates: (listener: (event: { updates: Uint8Array[]; inferences: [] }) => void) => {
    listeners.add(listener); return () => { listeners.delete(listener); };
  }, documentRead: mock(async () => ({ value: new Uint8Array([0]) })),
  documentReadAt: mock(async () => ({ status: 'ok', version: 'native-1', value: null })) } as unknown as ResidentEngineWorkerClient;
  const session = { submit: mock(() => ({ seq: 1, ack: Promise.resolve({ seq: 1, status: 'applied' }) })) } as unknown as DocxWorkerEditSession;
  const viewStore = createPeerlessViewStore({ identity: { session: 'one', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: { anchor: { story: 'body', paraId: 'p', offset: 0, sticky: new Uint8Array([1]) },
      focus: { story: 'body', paraId: 'p', offset: 0, sticky: new Uint8Array([1]) } }, anchors: null, editState: undefined, geometry: null });
  const binding = { session, viewStore } as PeerlessEditSessionBinding;
  return { client, session, binding, viewStore, listeners };
}

test('replica is published only after ready, rebinds handlers, and retires before null', async () => {
  const s = setup();
  const first = mock<(replica: CollaborationReplica | null) => void>(() => {});
  const next = mock<(replica: CollaborationReplica | null) => void>(() => {});
  let lifecycle: ReturnType<typeof usePeerlessCollaboration>;
  function Harness({ options, bytes = source }: { options: DocxEditorCollaborationOptions; bytes?: Uint8Array }) {
    lifecycle = usePeerlessCollaboration(options, bytes);
    return null;
  }
  const view = render(<Harness options={{ clientId: 41, initialUpdate: source, onReplica: first }} />);
  expect(first).not.toHaveBeenCalled();
  expect(lifecycle!.openOptions).toEqual({ clientId: 41, initialUpdate: source });
  await act(async () => lifecycle!.ready(s.client, s.binding));
  const adapter = first.mock.calls[0]![0]!;
  view.rerender(<Harness options={{ clientId: 41, initialUpdate: source, onReplica: next }} />);
  expect(first.mock.calls.at(-1)![0]).toBeNull();
  expect(next.mock.calls[0]![0]).toBe(adapter);
  view.rerender(<Harness options={{ clientId: 42, initialUpdate: source, onReplica: next }} />);
  await expect(adapter.applyUpdate(source)).rejects.toThrow('retired');
  expect(next.mock.calls.at(-1)![0]).toBeNull();
  await act(async () => lifecycle!.ready(s.client, s.binding));
  expect(next.mock.calls.at(-1)![0]?.clientId).toBe(42);
  const second = next.mock.calls.at(-1)![0]!;
  view.rerender(<Harness options={{ clientId: 42, initialUpdate: source, onReplica: next }} bytes={new Uint8Array([2])} />);
  await expect(second.applyUpdate(source)).rejects.toThrow('retired');
  await act(async () => lifecycle!.ready(s.client, s.binding));
  view.unmount();
  expect(next.mock.calls.at(-1)![0]).toBeNull();
});

test('async replica uses the edit session fence and forwards local updates', async () => {
  const s = setup();
  const replica = createPeerlessCollaborationReplica(s.client, s.session, 41);
  const listener = mock(() => {});
  replica.onUpdate(listener);
  for (const receive of s.listeners) receive({ updates: [source], inferences: [] });
  expect(listener).toHaveBeenCalledWith(source, 'local');
  await replica.async.applyUpdate(source);
  expect(s.session.submit).toHaveBeenCalledWith('applyRemoteUpdate', { update: source.buffer }, { lane: 'collab', fence: true });
  replica.retire();
  expect(s.listeners.size).toBe(0);
  await expect(replica.async.encodeStateVector()).rejects.toThrow('retired');
});

test('local cursor publication suppresses broadcast on document changes and clears on unmount', async () => {
  const s = setup();
  const presence: CollaborationPresence = { peers: [], setCursor: mock(() => {}), onPeers: () => () => {} };
  const view = render(<PeerlessRemotePresence binding={s.binding} client={s.client} presence={presence} pages={[0, 1]} />);
  await waitFor(() => expect(presence.setCursor).toHaveBeenCalledWith({ story: 'body', anchor: new Uint8Array([1]), head: new Uint8Array([1]) }, true));
  await act(async () => s.viewStore.acceptWorkerState(s.viewStore.getSnapshot().identity, { epoch: 1, version: 1, throughSeq: 1,
    selection: { anchor: { story: 'body', paraId: 'p', offset: 1, sticky: new Uint8Array([2]) }, focus: { story: 'body', paraId: 'p', offset: 1, sticky: new Uint8Array([2]) } } }));
  expect(presence.setCursor).toHaveBeenLastCalledWith({ story: 'body', anchor: new Uint8Array([2]), head: new Uint8Array([2]) }, false);
  view.unmount();
  expect(presence.setCursor).toHaveBeenLastCalledWith(null);
});

test('presence overlay ignores a reply from an older installed geometry', async () => {
  const s = setup();
  let resolveRead: (value: unknown) => void;
  s.client.documentReadAt = mock(() => new Promise(resolve => { resolveRead = resolve; })) as ResidentEngineWorkerClient['documentReadAt'];
  const identity = s.viewStore.getSnapshot().identity;
  const geometry = { identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, documentVersion: 'native-1',
    pages: [{ page: { pageId: '1', pageIndex: 0 }, bounds: { x: 0, y: 0, width: 200, height: 300 }, transform: [1, 0, 0, 1, 0, 0] as const }], caret: null, selectionRects: [] };
  s.viewStore.installFrameGeometry(geometry);
  const presence: CollaborationPresence = { peers: [{ clientId: 42, clock: 1, user: { name: 'Writer', color: '#f00' }, cursor: null, inferredCursor: null, cursorMovedAt: 1, lastSeenAt: 1 }],
    setCursor() {}, onPeers: () => () => {} };
  const view = render(<PeerlessRemotePresence binding={s.binding} client={s.client} presence={presence} pages={[0, 1]} />);
  await waitFor(() => expect(s.client.documentReadAt).toHaveBeenCalled());
  await act(async () => {
    s.viewStore.installFrameGeometry({ ...geometry, frameEpoch: 2 });
    resolveRead!({ status: 'ok', version: 'native-1', value: { version: 'native-1', frameEpoch: 1, layoutRevision: 1,
      peers: [{ clientId: 42, clock: 1, status: 'ok', rects: [], caret: { pageId: '1', pageIndex: 0, x: 12, y: 10, width: 0, height: 14 } }] } });
  });
  expect(view.queryByTestId('remote-presence-caret')).toBeNull();
});

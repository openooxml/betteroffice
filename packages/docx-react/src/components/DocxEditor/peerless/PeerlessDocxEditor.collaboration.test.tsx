import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import { useEffect, useLayoutEffect, useState } from 'react';
import * as yrs from '@betteroffice/docx/yrs';
import * as peerless from '@betteroffice/docx/peerless';
import * as measurement from '@betteroffice/docx/layout';
import * as paged from './PeerlessPagedEditor';
import { PeerlessDocxEditor as DocxEditor } from './PeerlessDocxEditor';
import { setMissingWorkerCapabilitiesForTests } from '../internals/engineChoice';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';
import { useWorkerCollaboration } from './useWorkerCollaboration';
import { PeerlessRemotePresence } from './PeerlessRemotePresence';
import { createPeerlessViewStore } from './viewStore';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { CollaborationPresence, CollaborationReplica } from '@betteroffice/docx/collaboration';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { DocxEditorCollaborationOptions } from '../types';
import type { Document } from '@betteroffice/docx/types/document';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
const restores: Array<() => void> = [];
afterEach(() => { cleanup(); for (const restore of restores.splice(0).reverse()) restore(); });
afterAll(async () => {
  await act(async () => { await new Promise<void>(resolve => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});
const source = new Uint8Array([1]);
type CollaborationUpdateEvent = Parameters<Parameters<ResidentEngineWorkerClient['onCollabUpdates']>[0]>[0];

function setup() {
  const listeners = new Set<(event: CollaborationUpdateEvent) => void>();
  const client = { onCollabUpdates: (listener: (event: CollaborationUpdateEvent) => void) => {
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
  let lifecycle: ReturnType<typeof useWorkerCollaboration>;
  function Harness({ options, bytes = source }: { options: DocxEditorCollaborationOptions; bytes?: Uint8Array }) {
    lifecycle = useWorkerCollaboration(options, bytes);
    return null;
  }
  const view = render(<Harness options={{ clientId: 41, initialUpdate: source, onReplica: first }} />);
  expect(first).not.toHaveBeenCalled();
  expect(lifecycle!.openOptions).toEqual({ clientId: 41, initialUpdate: source });
  await act(async () => { lifecycle!.begin(s.client); lifecycle!.ready(s.client, s.binding); });
  const adapter = first.mock.calls[0]![0]!;
  view.rerender(<Harness options={{ clientId: 41, initialUpdate: source, onReplica: next }} />);
  expect(first.mock.calls.at(-1)![0]).toBeNull();
  expect(next.mock.calls[0]![0]).toBe(adapter);
  view.rerender(<Harness options={{ clientId: 42, initialUpdate: source, onReplica: next }} />);
  await expect(adapter.applyUpdate(source)).rejects.toThrow('retired');
  expect(next.mock.calls.at(-1)![0]).toBeNull();
  await act(async () => { lifecycle!.begin(s.client); lifecycle!.ready(s.client, s.binding); });
  expect(next.mock.calls.at(-1)![0]?.clientId).toBe(42);
  const second = next.mock.calls.at(-1)![0]!;
  view.rerender(<Harness options={{ clientId: 42, initialUpdate: source, onReplica: next }} bytes={new Uint8Array([2])} />);
  await expect(second.applyUpdate(source)).rejects.toThrow('retired');
  await act(async () => { lifecycle!.begin(s.client); lifecycle!.ready(s.client, s.binding); });
  view.unmount();
  expect(next.mock.calls.at(-1)![0]).toBeNull();
});

test('async replica uses the edit session fence and forwards local updates', async () => {
  const s = setup();
  const submit = s.session.submit;
  const replica = createPeerlessCollaborationReplica(s.client, s.session, 41);
  const listener = mock(() => {});
  replica.onUpdate(listener);
  for (const receive of s.listeners) receive({ updates: [source], inferences: [] });
  expect(listener).toHaveBeenCalledWith(source, 'local');
  await replica.async.applyUpdate(source);
  expect(submit).toHaveBeenCalledWith('applyRemoteUpdate', { update: source.buffer }, { lane: 'collab', fence: true });
  replica.retire();
  expect(s.listeners.size).toBe(0);
  await expect(replica.async.encodeStateVector()).rejects.toThrow('retired');
});

test('local cursor publication suppresses broadcast on document changes and clears on unmount', async () => {
  const s = setup();
  const presence: CollaborationPresence = { peers: [], setCursor: mock(() => {}), onPeers: () => () => {} };
  const view = render(<PeerlessRemotePresence binding={s.binding} client={s.client} presence={presence} pages={[0, 1]} />);
  await waitFor(() => expect(presence.setCursor).toHaveBeenCalledWith({ story: 'body', anchor: new Uint8Array([1]), head: new Uint8Array([1]) }, true));
  await act(async () => s.viewStore.acceptWorkerState(s.viewStore.getSnapshot().identity, { epoch: 1, version: 1, throughSeq: 1, presentedThroughSeq: 1, editState: undefined,
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
  view.rerender(<PeerlessRemotePresence binding={s.binding} client={s.client} presence={presence} pages={[1, 2]} />);
  expect(s.client.documentReadAt).toHaveBeenCalledTimes(1);
  await act(async () => {
    s.viewStore.installFrameGeometry({ ...geometry, frameEpoch: 2 });
    resolveRead!({ status: 'ok', version: 'native-1', value: { version: 'native-1', frameEpoch: 1, layoutRevision: 1,
      peers: [{ clientId: 42, clock: 1, status: 'ok', rects: [], caret: { pageId: '1', pageIndex: 0, x: 12, y: 10, width: 0, height: 14 } }] } });
  });
  await waitFor(() => expect(s.client.documentReadAt).toHaveBeenCalledTimes(2));
  expect(view.queryByTestId('remote-presence-caret')).toBeNull();
});

test('presence reads an unresolved peer again once the document version changes', async () => {
  const s = setup();
  let status: 'ok' | 'unresolved' = 'unresolved';
  const reads: number[][] = [];
  s.client.documentReadAt = mock(async (read: Extract<Parameters<ResidentEngineWorkerClient['documentReadAt']>[0], { kind: 'collabPresence' }>) => {
    reads.push(read.peers.map(peer => peer.clientId));
    return { status: 'ok', version: read.expectVersion, value: { version: read.expectVersion, frameEpoch: read.expectFrameEpoch,
      layoutRevision: read.expectLayoutRevision, peers: read.peers.map(peer => ({ clientId: peer.clientId, clock: peer.clock, status, rects: [],
        caret: status === 'ok' ? { pageId: '1', pageIndex: 0, x: 12, y: 10, width: 0, height: 14 } : null })) } };
  }) as unknown as ResidentEngineWorkerClient['documentReadAt'];
  const identity = s.viewStore.getSnapshot().identity;
  const geometry = { identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, documentVersion: 'native-1',
    pages: [{ page: { pageId: '1', pageIndex: 0 }, bounds: { x: 0, y: 0, width: 200, height: 300 }, transform: [1, 0, 0, 1, 0, 0] as const }], caret: null, selectionRects: [] };
  s.viewStore.installFrameGeometry(geometry);
  const presence: CollaborationPresence = { peers: [{ clientId: 42, clock: 1, user: { name: 'Writer', color: '#f00' }, cursor: { story: 'body', anchor: new Uint8Array([9]), head: new Uint8Array([9]) },
    inferredCursor: null, cursorMovedAt: 1, lastSeenAt: 1 }], setCursor() {}, onPeers: () => () => {} };
  const view = render(<PeerlessRemotePresence binding={s.binding} client={s.client} presence={presence} pages={[0, 1]} />);
  await waitFor(() => expect(reads).toEqual([[42]]));
  await act(async () => s.viewStore.installFrameGeometry({ ...geometry, frameEpoch: 2 }));
  await waitFor(() => expect(reads.at(-1)).toEqual([]));
  status = 'ok';
  await act(async () => s.viewStore.installFrameGeometry({ ...geometry, frameEpoch: 3, documentVersion: 'native-2' }));
  await waitFor(() => expect(reads.at(-1)).toEqual([42]));
  await waitFor(() => expect(view.getByTestId('remote-presence-caret').dataset.clientId).toBe('42'));
});

const hostDocument = () => ({ package: {
  document: { content: [], sections: [] }, theme: {}, numbering: {}, settings: {}, fontTable: { fonts: [] },
  relationshipEntries: [], mediaEntries: [], chartEntries: [],
} }) as unknown as Document;

function mockWorkerEditor() {
  setMissingWorkerCapabilitiesForTests([]);
  const worker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  Object.defineProperty(globalThis, 'Worker', { configurable: true, value: class {
    postMessage() {} terminate() {}
  } });
  restores.push(() => { if (worker) Object.defineProperty(globalThis, 'Worker', worker); else Reflect.deleteProperty(globalThis, 'Worker'); });
  const open = spyOn(peerless.ResidentSourceWorkerClient.prototype, 'openSource').mockResolvedValue({ document: hostDocument(), stateVector: source });
  const prewarm = spyOn(peerless.ResidentSourceWorkerClient.prototype, 'prewarm').mockResolvedValue();
  const requirements = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'fontRequirements').mockResolvedValue('[]');
  const textEngine = spyOn(measurement, 'getRustTextEngine').mockResolvedValue({ registerFont: () => 0,
    registerSubstituteFont: () => 0, clearFonts() {} });
  const init = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'commandInit').mockImplementation(async request => ({
    session: request.session, incarnation: 1, epoch: 1, version: 0, throughSeq: 0, outcomes: [], selection: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
    frame: { coversThroughSeq: 0, revision: 1, payload: { frame: new ArrayBuffer(0), caret: { frameEpoch: 1, caretRect: null }, layoutRevision: 1 } },
  }));
  const connection = spyOn(peerless, 'createResidentCommandConnection').mockImplementation(async (client, initial) => ({
    initial, session: 'one', incarnation: 1, epoch: 1, client: { dispose: async () => client.destroy() },
  }) as unknown as peerless.ResidentCommandConnection);
  const surface = spyOn(paged, 'PeerlessPagedEditor').mockImplementation(function Surface({ sessionOptions, onReady }) {
    const [binding, setBinding] = useState<PeerlessEditSessionBinding | null>(null);
    useEffect(() => {
      const controller = new AbortController();
      void sessionOptions.openConnection(sessionOptions.source, controller.signal).then(() => {
        if (controller.signal.aborted) return;
        const s = setup();
        Object.assign(s.session, { state: { editState: { readOnly: false }, throughSeq: 0 }, dispose() {} });
        Object.assign(s.binding, { admission: { setSuggesting() {} } });
        setBinding(s.binding);
      }, sessionOptions.onError);
      return () => controller.abort();
    }, [sessionOptions.source]);
    useLayoutEffect(() => { if (binding) onReady?.(binding); }, [binding, onReady]);
    return null;
  });
  restores.push(() => open.mockRestore(), () => prewarm.mockRestore(), () => requirements.mockRestore(), () => textEngine.mockRestore(),
    () => init.mockRestore(), () => connection.mockRestore(), () => surface.mockRestore());
  return open;
}

test('editor opens the requested room, reopens on room and document changes, and only rebinds callbacks', async () => {
  const open = mockWorkerEditor();
  const first = mock<(replica: CollaborationReplica | null) => void>(() => {});
  const next = mock<(replica: CollaborationReplica | null) => void>(() => {});
  const initialUpdate = new Uint8Array([2]);
  const view = render(<DocxEditor experimentalPeerless documentBuffer={source} showToolbar={false} collaboration={{ clientId: 41, initialUpdate, onReplica: first }} />);
  await waitFor(() => expect(first).toHaveBeenCalledTimes(1));
  expect(open.mock.calls[0]![0]?.collaboration).toEqual({ clientId: 41, initialUpdate });
  const original = first.mock.calls[0]![0]!;
  view.rerender(<DocxEditor experimentalPeerless documentBuffer={source} showToolbar={false} collaboration={{ clientId: 41, initialUpdate, onReplica: next }} />);
  await waitFor(() => expect(next).toHaveBeenCalledWith(original));
  expect(open).toHaveBeenCalledTimes(1);
  view.rerender(<DocxEditor experimentalPeerless documentBuffer={source} showToolbar={false} collaboration={{ clientId: 42, initialUpdate, onReplica: next }} />);
  await waitFor(() => expect(next.mock.calls.at(-1)![0]?.clientId).toBe(42));
  expect(open).toHaveBeenCalledTimes(2);
  await expect(original.applyUpdate(source)).rejects.toThrow('retired');
  const second = next.mock.calls.at(-1)![0]!;
  const replacement = new Uint8Array([3]);
  view.rerender(<DocxEditor experimentalPeerless documentBuffer={replacement} showToolbar={false} collaboration={{ clientId: 42, initialUpdate, onReplica: next }} />);
  await waitFor(() => expect(open).toHaveBeenCalledTimes(3));
  await waitFor(() => expect(next.mock.calls.at(-1)![0]).not.toBe(second));
  await expect(second.applyUpdate(source)).rejects.toThrow('retired');
  const swapped = next.mock.calls.at(-1)![0]!;
  view.rerender(<DocxEditor experimentalPeerless documentBuffer={replacement} showToolbar={false} />);
  await waitFor(() => expect(open).toHaveBeenCalledTimes(4));
  expect(open.mock.calls[3]![0]?.collaboration).toBeUndefined();
  expect(next.mock.calls.at(-1)![0]).toBeNull();
  await expect(swapped.applyUpdate(source)).rejects.toThrow('retired');
  view.rerender(<DocxEditor experimentalPeerless documentBuffer={replacement} showToolbar={false} collaboration={{ clientId: 42, initialUpdate, onReplica: next }} />);
  await waitFor(() => expect(open).toHaveBeenCalledTimes(5));
  await waitFor(() => expect(next.mock.calls.at(-1)![0]?.clientId).toBe(42));
  view.unmount();
  expect(next.mock.calls.at(-1)![0]).toBeNull();
});

test('a parsed document reopens its worker on a room change', async () => {
  const open = mockWorkerEditor();
  const document = hostDocument();
  const onReplica = mock<(replica: CollaborationReplica | null) => void>(() => {});
  const view = render(<DocxEditor experimentalPeerless document={document} showToolbar={false} collaboration={{ clientId: 41, onReplica }} />);
  await waitFor(() => expect(onReplica).toHaveBeenCalledTimes(1));
  view.rerender(<DocxEditor experimentalPeerless document={document} showToolbar={false} collaboration={{ clientId: 42, onReplica }} />);
  await waitFor(() => expect(onReplica.mock.calls.at(-1)![0]?.clientId).toBe(42));
  expect(open).toHaveBeenCalledTimes(2);
  expect(open.mock.calls[1]![0]?.collaboration?.clientId).toBe(42);
  view.unmount();
  expect(onReplica.mock.calls.at(-1)![0]).toBeNull();
});

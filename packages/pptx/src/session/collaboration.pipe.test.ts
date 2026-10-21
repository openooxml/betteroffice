import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { attachCollaborationPipe, collaborationPipeFor } from '../collaboration/asyncReplica';
import { CollaborationProvider } from '../collaboration/provider';
import { decodeAwarenessUpdate, decodeMessages, encodeAwarenessUpdate, encodeSyncStep2, encodeUpdate } from '../collaboration/protocol';
import type { AsyncCollaborationReplica, CollaborationError, CollaborationProviderOptions, CollaborationTransport, CollaborationTransportEvent } from '../collaboration/types';
import type { PptxReadResult } from '../edits';
import type { SlideDisplayList } from '../types';
import { initWasm, openPresentation, type PresentationHandle } from '../wasm/loader';
import { createPresentationCollaboration } from './collaboration';
import { createPresentationCommandClient, createSharedPresentationCommandClient } from './commandClient';
import { presentationCollaborationPipes } from './commandCollaboration';
import type { PresentationCommandAck, PresentationCommandEnvelope, PresentationCommandIntent } from './commandProtocol';
import { createPresentationSessionHost } from './host';

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns} xmlns:r="${rel}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>abc</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

let bytes: Uint8Array;
let font: Uint8Array;
beforeAll(async () => {
  bytes = await deck();
  font = new Uint8Array(await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
});

class Wire implements CollaborationTransport {
  readonly frames: Uint8Array[] = [];
  readonly calls: string[] = [];
  readonly listeners = new Set<(event: CollaborationTransportEvent) => void>();
  accept = true;
  connect() { this.calls.push('connect'); this.emit({ type: 'open' }); }
  disconnect() { this.calls.push('disconnect'); }
  send(data: Uint8Array) { this.frames.push(data.slice()); return this.accept; }
  onEvent(listener: (event: CollaborationTransportEvent) => void) {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }
  emit(event: CollaborationTransportEvent) { for (const listener of [...this.listeners]) listener(event); }
}

async function until(done: () => boolean) {
  for (let turn = 0; turn < 1000; turn++) {
    if (done()) return;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error('Timed out waiting for collaboration');
}

async function open(options: CollaborationProviderOptions = {}, shared = true) {
  const pair = createInProcessPair();
  const requests: PresentationCommandEnvelope[] = [];
  const transferred: ArrayBuffer[][] = [];
  const transport: SessionTransport = { ...pair.client, post(message, transfer) {
    if (isClientMessage(message) && message.kind === 'call' && ['edit', 'command'].includes(message.method)) {
      requests.push(structuredClone(message.args[0] as PresentationCommandEnvelope));
      transferred.push((transfer ?? []).filter((buffer): buffer is ArrayBuffer => buffer instanceof ArrayBuffer));
    }
    pair.client.post(message, transfer);
  } };
  let disposeHost = () => {};
  const replies: unknown[] = [];
  createPresentationSessionHost({ ...pair.host, post(message, transfer) {
    if (isHostMessage(message) && message.kind === 'reply' && 'collabPipe' in message) replies.push(structuredClone(message));
    pair.host.post(message, transfer);
  }, listen(listener) {
    disposeHost = () => listener({ protocol: 1, kind: 'dispose' });
    return pair.host.listen(listener);
  } }, { initWasm: async () => {} });
  const client = await (shared ? createSharedPresentationCommandClient : createPresentationCommandClient)(bytes,
    { session: 'pipe', incarnation: 'pipe-worker', collaboration: true, clientId: 8101,
      fonts: [{ family: 'Liberation Sans', bytes: font.slice().buffer }] }, transport);
  const unavailable = async (): Promise<never> => { throw new Error('Main must not read or apply the replica'); };
  const replica: AsyncCollaborationReplica = { clientId: 8101, encodeStateVector: unavailable,
    encodeStateAsUpdate: unavailable, applyUpdate: unavailable,
    onUpdate: () => { throw new Error('Main must not observe the replica'); } };
  const post = async (payload: Extract<PresentationCommandIntent, { kind: 'collabPipe' }>['payload']) => {
    const ack = await client.dispatch({ kind: 'collabPipe', payload });
    if (ack.outcomes.some((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled')) {
      throw new Error('Pipe command refused');
    }
  };
  attachCollaborationPipe(replica, presentationCollaborationPipes.get(client)!.channel((payload) =>
    post({ ...payload, inputTime: performance.now(), undoGroup: 'remote' })));
  const wire = new Wire();
  const provider = new CollaborationProvider(replica, wire, options);
  const errors: CollaborationError[] = [];
  provider.onError((error) => errors.push(error));
  provider.connect();
  await until(() => provider.status === 'connected' && wire.frames.length > 0);
  await client.flush();
  return { client, replica, requests, transferred, replies, wire, provider, errors,
    replay(message: unknown) { pair.host.post(message); },
    fail() { pair.host.post({ protocol: 1, kind: 'failure', code: 'crash', message: 'Worker lost' }); },
    async dispose() { provider.destroy(); try { await client.dispose(); } finally { disposeHost(); } },
  };
}

function content(result: PptxReadResult) {
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

async function sync(room: Awaited<ReturnType<typeof open>>, peer: ReturnType<typeof openPresentation>) {
  room.wire.emit({ type: 'message', data: encodeSyncStep2(peer.encodeStateAsUpdate()) });
  await until(() => room.provider.synced);
  await room.client.flush();
}

for (const shared of [false, true]) test(`synced listeners dispatch against the installed ${shared ? 'shared' : 'command'} acknowledgement`, async () => {
  const room = await open({}, shared);
  const peer = openPresentation(bytes, { clientId: 8102 });
  try {
    const story = content(peer.readContent()).stories[0]!;
    peer.insertText(story.storyId, 0, 'remote ');
    const before = room.client.state.version;
    let installed: PresentationCommandAck = room.client.state;
    room.client.subscribe((ack) => { installed = ack; });
    let observed: { version: string; installedVersion: string; throughSeq: number; installedThroughSeq: number } | undefined;
    let operation: Promise<PresentationCommandAck> | undefined;
    room.provider.onStatus(({ synced }) => {
      if (!synced || operation) return;
      observed = { version: room.client.state.version, installedVersion: installed.version,
        throughSeq: room.client.state.throughSeq, installedThroughSeq: installed.throughSeq };
      operation = room.client.dispatch({ kind: 'invoke', payload: {
        method: 'setSlideNotes', args: [story.slideId, 'Synced notes'], inputTime: 1, undoGroup: 'notes',
      } });
    });
    room.wire.emit({ type: 'message', data: encodeSyncStep2(peer.encodeStateAsUpdate()) });
    await until(() => operation !== undefined);
    const ack = await operation!;
    expect(ack.outcomes).toMatchObject([{ status: 'applied', committed: true }]);
    expect(observed!.version).not.toBe(before);
    expect(observed!.installedVersion).toBe(observed!.version);
    expect(observed!.installedThroughSeq).toBe(observed!.throughSeq);
    expect(content((await room.client.read()).result).stories[0]!.text).toBe(`remote ${story.text}`);
    expect(room.errors).toEqual([]);
  } finally { peer.dispose(); await room.dispose(); }
});

for (const shared of [false, true]) test(`a replayed ${shared ? 'shared' : 'command'} reply leaves later pipe output flowing`, async () => {
  const room = await open({}, shared);
  const peer = openPresentation(bytes, { clientId: 8103 });
  try {
    const answered = room.replies.find((reply) => {
      const value = (reply as { ok?: boolean; value?: unknown }).value;
      return value !== null && typeof value === 'object' && 'outcomes' in value;
    });
    expect(answered).toBeDefined();
    room.replay(answered);
    const story = content(peer.readContent()).stories[0]!;
    peer.insertText(story.storyId, 0, 'remote ');
    await sync(room, peer);
    expect(content((await room.client.read()).result).stories[0]!.text).toBe(`remote ${story.text}`);
    expect(room.errors).toEqual([]);
  } finally { peer.dispose(); await room.dispose(); }
});

for (const kind of ['cursor', 'awareness'] as const) test(`non-mutating pipe ${kind} preserves a typing undo group`, async () => {
  const room = await open();
  try {
    const story = content((await room.client.read()).result).stories[0]!;
    const first = await room.client.dispatch({ kind: 'insertText', payload: {
      inputTime: 100, undoGroup: 'typing', text: 'X', at: { source: 'confirmed', selection: {
        slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
        anchor: 0, focus: 0, affinity: 'downstream' } },
    } });
    expect(first.outcomes).toMatchObject([{ status: 'applied', committed: true }]);
    const before = room.requests.length;
    if (kind === 'cursor') room.provider.setCursor({ slideId: story.slideId });
    else room.wire.emit({ type: 'message', data: encodeAwarenessUpdate([{ clientId: 8102, clock: 1, state: null }]) });
    await until(() => room.requests.length > before);
    await room.client.flush();
    expect(room.client.state.version).toBe(first.version);
    expect(room.client.state.selection).toEqual(first.selection);
    const second = await room.client.dispatch({ kind: 'insertText', payload: {
      inputTime: 101, undoGroup: 'typing', text: 'Y', at: { source: 'confirmed',
        selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
          anchor: 1, focus: 1, affinity: 'downstream' } },
    } });
    expect(second.outcomes).toMatchObject([{ status: 'applied', committed: true }]);
    expect(content((await room.client.read()).result).stories[0]!.text).toBe('XYabc');
    const undone = await room.client.dispatch({ kind: 'undo', payload: { inputTime: 102, undoGroup: 'undo' } });
    expect(undone.outcomes).toMatchObject([{ status: 'applied', committed: true }]);
    expect(content((await room.client.read()).result).stories[0]!.text).toBe('abc');
    expect(undone.editState.canUndo).toBe(false);
    expect(room.errors).toEqual([]);
  } finally { await room.dispose(); }
});

test('cursor pipe commands preserve a word drag gesture', async () => {
  const room = await open();
  try {
    const story = content((await room.client.read()).result).stories[0]!;
    const frame = room.client.state.frame!;
    const list = JSON.parse(new TextDecoder().decode(frame.displayList)) as SlideDisplayList;
    const box = list.primitives.find((primitive) => primitive.kind === 'textBox');
    if (!box || box.kind !== 'textBox') throw new Error('Missing text box');
    const line = box.lines[0]!;
    const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = frame;
    const payload = { inputTime: 100, undoGroup: 'drag', granularity: 'word' as const, extend: false,
      frame: { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation },
      x: line.caretStops[1]!.x, y: line.y + line.height / 2,
      at: { source: 'confirmed' as const, selection: { slideId: story.slideId, shapeId: story.shapeId,
        storyId: story.storyId, anchor: 1, focus: 1, affinity: 'downstream' as const } } };
    const selected = await room.client.dispatch({ kind: 'selectAtPoint', payload });
    expect(selected.outcomes).toMatchObject([{ status: 'applied', committed: false }]);
    expect(selected.selection).toMatchObject({ anchor: 0, focus: 3 });
    const before = room.requests.length;
    room.provider.setCursor({ slideId: story.slideId });
    await until(() => room.requests.length > before);
    await room.client.flush();
    const dragged = await room.client.dispatch({ kind: 'selectAtPoint', payload: { ...payload, drag: true } });
    expect(dragged.outcomes).toMatchObject([{ status: 'applied', committed: false }]);
    expect(dragged.selection).toMatchObject({ anchor: 0, focus: 3 });
    expect(room.errors).toEqual([]);
  } finally { await room.dispose(); }
});

for (const shared of [true, false]) test(`worker burst merges once with ${shared ? 'shared' : 'direct'} commands`, async () => {
  const room = await open({}, shared);
  const peer = openPresentation(bytes, { clientId: 8102 });
  try {
    expect(collaborationPipeFor(room.replica)).toBeDefined();
    await sync(room, peer);
    const updates: Uint8Array[] = [];
    peer.onUpdate((update, origin) => { if (origin === 'local') updates.push(update); });
    const story = content((await room.client.read()).result).stories[0]!;
    for (let index = 0; index < 24; index++) peer.insertText(story.storyId, 0, `${index} `);
    expect(updates).toHaveLength(24);
    const before = { commands: room.requests.length, version: room.client.state.version, frame: room.client.state.frame!.epoch };
    const acks: PresentationCommandAck[] = [];
    room.client.subscribe((ack) => acks.push(ack));
    for (const update of updates) room.wire.emit({ type: 'message', data: encodeUpdate(update) });
    await until(() => acks.length > 0);
    const commands = room.requests.slice(before.commands);
    expect(commands.map((command) => command.kind)).toEqual(['collabPipe']);
    const command = commands[0]!;
    expect(command.kind === 'collabPipe' && command.payload.ops.filter((op) => op.op === 'event')).toHaveLength(24);
    expect(room.transferred[room.transferred.length - 1]).toHaveLength(24);
    expect(room.transferred[room.transferred.length - 1].every((buffer) => buffer.byteLength === 0)).toBe(true);
    expect(acks).toHaveLength(1);
    expect(acks[0].frame?.epoch).toBe(before.frame + 1);
    expect(acks[0].version).not.toBe(before.version);
    expect(acks[0].outcomes).toMatchObject([{ status: 'applied', committed: true }]);
    expect(room.client.state.editState.canUndo).toBe(false);
    const converged = content((await room.client.read()).result).stories[0]!;
    expect(converged.text).toBe(content(peer.readContent()).stories[0]!.text);
    await room.client.dispatch({ kind: 'insertText', payload: {
      inputTime: 1, undoGroup: 'typing', text: 'local ', at: { source: 'confirmed', selection: {
        slideId: converged.slideId, shapeId: converged.shapeId, storyId: converged.storyId,
        anchor: 0, focus: 0, affinity: 'downstream' } },
    } });
    await room.client.dispatch({ kind: 'undo', payload: { inputTime: 2, undoGroup: 'undo' } });
    expect(content((await room.client.read()).result).stories[0]!.text).toBe(converged.text);
    expect(room.client.state.editState.canUndo).toBe(false);
    expect(room.errors).toEqual([]);
  } finally { peer.dispose(); await room.dispose(); }
});

test('presence reaches main and timed cursor output is posted without another command', async () => {
  const room = await open({ user: { name: 'Ada' } });
  const peer = openPresentation(bytes, { clientId: 8102 });
  try {
    await sync(room, peer);
    const seen: number[] = [];
    room.provider.onPresence((peers) => seen.push(peers.length));
    const cursor = { slideId: room.client.state.activeSlideId!, shapeId: 'shape' };
    room.wire.emit({ type: 'message', data: encodeAwarenessUpdate([{ clientId: 8102, clock: 1,
      state: { clientId: 8102, clock: 1, user: { name: 'Bo', color: '#0B57D0' }, cursor } }]) });
    await until(() => room.provider.peers.length === 1);
    expect(room.provider.peers[0].state).toMatchObject({ clientId: 8102, user: { name: 'Bo' }, cursor });
    expect(seen).toEqual([0, 1]);
    const awareness = () => room.wire.frames.flatMap((frame) => decodeMessages(frame))
      .flatMap((message) => message.type === 'awareness' ? decodeAwarenessUpdate(message.update) : []);
    room.provider.setCursor(cursor);
    await until(() => awareness().some((entry) => entry.state?.cursor?.shapeId === 'shape'));
    const next = { ...cursor, shapeId: 'next' };
    const beforeCursor = room.requests.length;
    room.provider.setCursor(next);
    await until(() => room.requests.length > beforeCursor);
    await room.client.flush();
    const requests = room.requests.length;
    await until(() => awareness().some((entry) => entry.state?.cursor?.shapeId === 'next'));
    expect(room.requests).toHaveLength(requests);
    room.provider.destroy();
    const states = awareness();
    expect(states[states.length - 1]?.state).toBeNull();
    expect(room.provider.status).toBe('destroyed');
    expect(room.wire.listeners.size).toBe(0);
    expect(room.errors).toEqual([]);
  } finally { peer.dispose(); await room.dispose(); }
});

for (const prefix of [false, true]) test(`a refused update ${prefix ? 'keeps the accepted prefix' : 'leaves the session unchanged'}`, async () => {
  const room = await open();
  const peer = openPresentation(bytes, { clientId: 8102 });
  try {
    await sync(room, peer);
    const story = content((await room.client.read()).result).stories[0]!;
    const before = room.client.state.version;
    const vector = peer.encodeStateVector();
    peer.insertText(story.storyId, 0, 'accepted ');
    if (prefix) room.wire.emit({ type: 'message', data: encodeUpdate(peer.encodeStateAsUpdate(vector)) });
    room.wire.emit({ type: 'message', data: encodeUpdate(Uint8Array.of(0xff)) });
    await until(() => room.errors.length > 0);
    expect(room.errors).toHaveLength(1);
    expect(room.errors[0]).toMatchObject({ code: 'replica' });
    expect(room.errors[0].message).toContain('Failed to apply remote update');
    expect(room.provider.status).toBe('disconnected');
    expect(room.provider.synced).toBe(false);
    expect(room.wire.listeners.size).toBe(0);
    const current = content((await room.client.read()).result).stories[0]!;
    expect(current.text).toBe(prefix ? `accepted ${story.text}` : story.text);
    if (!prefix) expect(room.client.state.version).toBe(before);
    expect((await room.client.dispatch({ kind: 'insertText', payload: {
      inputTime: 1, undoGroup: 'typing', text: 'local ', at: { source: 'confirmed', selection: {
        slideId: current.slideId, shapeId: current.shapeId, storyId: current.storyId,
        anchor: 0, focus: 0, affinity: 'downstream' } },
    } })).outcomes).toMatchObject([{ status: 'applied' }]);
    expect(room.client.state.editState.canUndo).toBe(true);
  } finally { peer.dispose(); await room.dispose(); }
});

test('disconnect and reconnect resync through the worker; destroy is final', async () => {
  const room = await open();
  const peer = openPresentation(bytes, { clientId: 8102 });
  try {
    await sync(room, peer);
    room.provider.disconnect();
    expect(room.provider.status).toBe('disconnected');
    expect(room.provider.synced).toBe(false);
    expect(room.wire.listeners.size).toBe(0);
    const story = content(peer.readContent()).stories[0]!;
    peer.insertText(story.storyId, 0, 'offline ');
    room.provider.connect();
    expect(room.provider.status).toBe('connecting');
    await until(() => room.provider.status === 'connected');
    await sync(room, peer);
    expect(content((await room.client.read()).result).stories[0]!.text).toBe(`offline ${story.text}`);
    room.provider.destroy();
    room.provider.connect();
    expect(room.provider.status).toBe('destroyed');
    expect(room.wire.calls).toEqual(['connect', 'disconnect', 'connect', 'disconnect']);
    expect(room.errors).toEqual([]);
  } finally { peer.dispose(); await room.dispose(); }
});

test('backpressure queues on main and drain retries', async () => {
  const room = await open();
  try {
    room.wire.accept = false;
    room.provider.setCursor({ slideId: room.client.state.activeSlideId! });
    await until(() => room.provider.pendingBytes > 0);
    const frames = room.wire.frames.length;
    room.wire.accept = true;
    room.wire.emit({ type: 'drain' });
    expect(room.provider.pendingBytes).toBe(0);
    expect(room.wire.frames.length).toBeGreaterThan(frames);
    expect(room.errors).toEqual([]);
  } finally { await room.dispose(); }
});

test('maxPendingBytes fails only the connection', async () => {
  const room = await open({ maxPendingBytes: 0 });
  try {
    room.wire.accept = false;
    room.provider.setCursor({ slideId: room.client.state.activeSlideId! });
    await until(() => room.errors.length > 0);
    expect(room.errors[0]).toMatchObject({ code: 'backpressure' });
    expect(room.provider.status).toBe('disconnected');
    expect(room.provider.pendingBytes).toBe(0);
    expect(room.wire.listeners.size).toBe(0);
    await room.client.read();
  } finally { await room.dispose(); }
});

test('worker loss retires the facade and a replacement worker resyncs', async () => {
  const room = await open();
  let replacement: Awaited<ReturnType<typeof open>> | undefined;
  const peer = openPresentation(bytes, { clientId: 8102 });
  try {
    room.fail();
    await until(() => room.errors.length > 0);
    expect(room.errors[0]).toMatchObject({ code: 'replica' });
    expect(room.provider.status).toBe('disconnected');
    expect(room.provider.synced).toBe(false);
    expect(room.wire.listeners.size).toBe(0);
    const story = content(peer.readContent()).stories[0]!;
    peer.insertText(story.storyId, 0, 'resynced ');
    replacement = await open();
    await sync(replacement, peer);
    expect(content((await replacement.client.read()).result).stories[0]!.text).toBe(`resynced ${story.text}`);
    expect(replacement.errors).toEqual([]);
  } finally { peer.dispose(); await room.dispose(); await replacement?.dispose(); }
});

for (const mode of ['refused', 'trap', 'changed'] as const) test(`pipe merge ${mode} follows the version failure rule`, () => {
  let version = 'before';
  const opened = { clientId: 91, version: () => version,
    encodeStateVector: () => Uint8Array.of(0), encodeStateAsUpdate: () => Uint8Array.of(0, 0),
    onUpdate: () => () => {},
    applyUpdate: () => {
      if (mode === 'trap') throw new WebAssembly.RuntimeError('Unsafe merge');
      if (mode === 'changed') version = 'after';
      throw new Error('Unsafe merge');
    },
  } as unknown as PresentationHandle;
  const pipe = createPresentationCollaboration(opened);
  try {
    pipe.run({ provider: 1, ops: [{ op: 'create', seq: 1, options: {} }, { op: 'connect', seq: 2 },
      { op: 'event', seq: 3, event: { type: 'open' } }] });
    pipe.drain();
    const run = () => pipe.run({ provider: 1, ops: [{ op: 'event', seq: 4,
      event: { type: 'message', data: new Uint8Array(encodeUpdate(Uint8Array.of(0, 0))).buffer } }] });
    if (mode === 'refused') {
      expect(run).not.toThrow();
      expect(pipe.drain()?.batches.flatMap((batch) => batch.out)).toContainEqual(
        expect.objectContaining({ do: 'error', error: { code: 'replica', message: 'Failed to apply remote update: Unsafe merge' } }));
    } else expect(run).toThrow('Unsafe merge');
  } finally { pipe.dispose(); }
});

test('pipe payloads reject invalid cursors, frame sizes and worker options', async () => {
  const room = await open();
  try {
    const before = room.client.state.version;
    const ops: unknown[][] = [
      [{ op: 'create', seq: 1, options: { maxPendingBytes: 1 } }],
      [{ op: 'cursor', seq: 1, cursor: { slideId: '' }, broadcast: true }],
      [{ op: 'cursor', seq: 1, cursor: { slideId: 'slide', shapeId: 1 }, broadcast: true }],
      [{ op: 'cursor', seq: 1, cursor: { slideId: 's'.repeat(1025) }, broadcast: true }],
      [{ op: 'event', seq: 1, event: { type: 'message', data: new ArrayBuffer(16 * 1024 * 1024 + 1) } }],
    ];
    for (const invalid of ops) {
      const ack = await room.client.dispatch({ kind: 'collabPipe', payload: {
        provider: 99, ops: invalid, inputTime: 1, undoGroup: 'remote',
      } } as PresentationCommandIntent);
      expect(ack.outcomes).toMatchObject([{ status: 'refused', reason: 'invalid-payload', committed: false }]);
      expect(ack.version).toBe(before);
    }
    expect(room.provider.status).toBe('connected');
    expect(room.errors).toEqual([]);
  } finally { await room.dispose(); }
});

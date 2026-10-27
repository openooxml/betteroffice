import { expect, mock, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { isCommandAck, isCommandEnvelope } from '../../../../shared/office-session/commands/types';
import { isClientMessage, isHostMessage, type HostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createPptxWorkerSession } from '../../../pptx/src/session/commandSession';
import type {
  PresentationCommandAck, PresentationCommandFrame, PresentationCommandIntent, PresentationCommandSelection,
} from '../../../pptx/src/session/commandProtocol';
import { createPresentationSessionHost } from '../../../pptx/src/session/host';
import { initWasm } from '../../../pptx/src/wasm/loader';
import {
  createPeerlessViewStore, installCoveringFrame, peerlessInputReference,
  type PeerlessConfirmedFrame, type PeerlessViewPainters, type PeerlessViewSnapshot,
} from './peerlessViewStore';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns} xmlns:r="${rel}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>abc</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'));
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  zip.file('ppt/presentation.xml', presentation.replace('</p:sldIdLst>', '<p:sldId id="257" r:id="rId2"/></p:sldIdLst>'));
  const relationships = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
  zip.file('ppt/_rels/presentation.xml.rels', relationships.replace('</Relationships>', `<Relationship Id="rId2" Type="${rel}/slide" Target="slides/slide2.xml"/></Relationships>`));
  zip.file('ppt/slides/slide2.xml', await zip.file('ppt/slides/slide1.xml')!.async('string'));
  return zip.generateAsync({ type: 'uint8array' });
}

async function open(painters: Partial<PeerlessViewPainters> = {}, echoEnabled = false, onError?: (error: Error) => void) {
  const pair = createInProcessPair();
  let disposeHost = () => {};
  const hostTransport: SessionTransport = { ...pair.host, listen(listener) {
    disposeHost = () => listener({ protocol: 1, kind: 'dispose' });
    return pair.host.listen(listener);
  } };
  const calls: { id: number; method: string; envelope: unknown }[] = [];
  const methods: string[] = [];
  const held: Extract<HostMessage, { kind: 'reply' }>[] = [];
  const arrived = new Map<number, ReturnType<typeof deferred<void>>>();
  let deliver: ((message: unknown) => void) | undefined;
  let holding = false;
  let sourceReads = 0;
  const source = await deck();
  const sourceBuffer = source.buffer;
  Object.defineProperty(source, 'buffer', {
    get() {
      sourceReads += 1;
      return sourceBuffer;
    },
  });
  const events = new EventTarget();
  const transport: SessionTransport = { ...pair.client,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call') methods.push(message.method);
      if (isClientMessage(message) && message.kind === 'call' &&
        (message.method === 'command' || message.method === 'edit')) {
        calls.push({ id: message.id, method: message.method, envelope: message.args[0] });
      }
      pair.client.post(message, transfer);
    },
    listen(listener) {
      deliver = listener;
      return pair.client.listen((message) => {
        if (holding && isHostMessage(message) && message.kind === 'reply' && calls.some((call) => call.id === message.id)) {
          held.push(message);
          for (const [count, waiter] of arrived) if (held.length >= count) waiter.resolve();
        } else listener(message);
      });
    },
  };
  createPresentationSessionHost(hostTransport, { initWasm: async () => {
    await initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  const off = transport.listen((message) => events.dispatchEvent(new MessageEvent('message', { data: message })));
  const worker = {
    postMessage(message: unknown, transfer?: Transferable[]) { transport.post(message, transfer); },
    addEventListener: events.addEventListener.bind(events),
    removeEventListener: events.removeEventListener.bind(events),
    terminate() { disposeHost(); off(); transport.close(); },
  } as unknown as Worker;
  const workerFactory = mock(() => worker);
  const errors = mock((_error: Error) => {});
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const owner = createPptxWorkerSession(source, { worker: workerFactory, onError: errors, clientId: 7811,
    fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }],
  });
  const client = await owner.ready;
  const read = (await client.read()).result;
  expect(read.ok).toBe(true);
  if (!read.ok) throw new Error(read.failure.message);
  const story = read.stories[0]!;
  const selection: PresentationCommandSelection = { slideId: story.slideId, shapeId: story.shapeId,
    storyId: story.storyId, anchor: 1, focus: 1, affinity: 'downstream' };
  if (echoEnabled) {
    await client.dispatch({ kind: 'replaceText', payload: {
      at: { source: 'confirmed', selection: { ...selection, anchor: 0, focus: 3 } }, text: 'abc',
      style: { bold: false, italic: false, underline: 'none', fontSizePt: 24,
        color: '#000000', fontFamily: 'Liberation Sans' }, inputTime: 0, undoGroup: 'setup',
    } });
  }
  await client.dispatch({ kind: 'setSelection', payload: { selection, inputTime: 0, undoGroup: 'selection' } });
  if (echoEnabled) expect(client.state.editState.toolbar.echoRun).toBe('after');
  calls.length = 0;
  const baseSeq = client.state.throughSeq;
  const accepted: PeerlessConfirmedFrame[] = [];
  const paintFrame = mock((frame: PeerlessConfirmedFrame, view: PeerlessViewSnapshot) => {
    painters.paintFrame?.(frame, view);
    accepted.push(frame);
  });
  const store = createPeerlessViewStore(client, {
    paintFrame, paintEcho: painters.paintEcho ?? (() => {}), restoreLine: painters.restoreLine ?? (() => {}),
  }, { echoEnabled, onError });
  const dispatch = store.dispatch.bind(store);
  store.dispatch = (intent) => {
    const operation = dispatch(intent);
    void operation.catch(() => {});
    return operation;
  };
  return { owner, client, store, selection, baseSeq, calls, methods, accepted, workerFactory, errors,
    get sourceReads() { return sourceReads; },
    hold() { holding = true; },
    replyFrame(): PresentationCommandFrame | undefined {
      const reply = held[0];
      if (!reply?.ok) return undefined;
      if (isCommandAck(reply.value)) return reply.value.frame?.payload as PresentationCommandFrame | undefined;
      return (reply.value as PresentationCommandAck).frame ?? undefined;
    },
    waitForReplies(count: number) {
      if (held.length >= count) return Promise.resolve();
      const waiter = deferred<void>();
      arrived.set(count, waiter);
      return waiter.promise;
    },
    releaseOne() { const message = held.shift(); if (message) deliver?.(message); },
    release() { holding = false; for (const message of held.splice(0)) deliver?.(message); },
    crash() { events.dispatchEvent(Object.assign(new Event('error'), { message: 'worker stopped' })); },
    async dispose() { holding = false; store.dispose(); await owner.dispose(); },
  };
}

function insert(selection: PresentationCommandSelection, text: string, undoGroup = 'typing'): PresentationCommandIntent {
  return { kind: 'insertText', payload: { text, at: { source: 'confirmed', selection }, inputTime: 100, undoGroup } };
}

function sharedWire(value: unknown) {
  expect(isCommandEnvelope(value)).toBe(true);
  if (!isCommandEnvelope(value)) throw new Error('Expected shared envelope');
  return value;
}

test('rec_editor_uses_shared_tickets_without_a_second_sequence_allocator', async () => {
  const value = await open();
  try {
    value.hold();
    const first = value.store.dispatch(insert(value.selection, 'X'));
    const second = value.store.dispatch(insert(value.selection, 'Y'));
    const third = value.store.dispatch(insert(value.selection, 'Z'));
    expect(third).toBe(second);
    expect(value.calls).toHaveLength(1);
    expect(value.store.getSnapshot()).toMatchObject({ waitingInputCount: 1, echo: null,
      pending: [{ seq: value.baseSeq + 1, echoed: false }] });
    expect(value.store.getSnapshot().activeLine?.text).toBe('abc');
    await value.waitForReplies(1);
    value.releaseOne();
    await first;
    await value.waitForReplies(1);
    const sent = value.calls.map((call) => sharedWire(call.envelope));
    expect(value.calls.map((call) => call.method)).toEqual(['edit', 'edit']);
    expect(sent.map((envelope) => envelope.seq)).toEqual([value.baseSeq + 1, value.baseSeq + 2]);
    expect(sent[1]!.afterSeq).toBe(sent[0]!.seq);
    expect(sent[1]!.payload).toMatchObject({ text: 'YZ', at: { source: 'afterSeq', seq: sent[0]!.seq } });
    expect(sent[1]!.basis).toEqual(sent[0]!.basis);
    expect(sent[1]!.expectVersion ?? false).toBe(false);
    value.release();
    const acks = await Promise.all([first, second, third]);
    expect(acks.map((ack) => ack.outcomes[0]!.seq)).toEqual([sent[0]!.seq, sent[1]!.seq, sent[1]!.seq]);
    expect(acks[2]).toBe(acks[1]);
    expect(value.store.getSnapshot()).toMatchObject({ waitingInputCount: 0, pending: [], echo: null,
      presentedThroughSeq: sent[1]!.seq, selection: { ...value.selection, anchor: 4, focus: 4 } });
    expect(value.store.getSnapshot().activeLine?.text).toBe('aXYZbc');
  } finally { await value.dispose(); }
});

test('rec_unsent_dependencies_use_ticket_identity_and_send_time_seq', async () => {
  const value = await open();
  try {
    value.hold();
    const first = value.store.dispatch({ kind: 'replaceText', payload: {
      inputTime: 100, undoGroup: 'typing', at: { source: 'confirmed', selection: value.selection }, text: 'X',
    } });
    const independent = value.store.dispatch({ kind: 'setReadOnly', payload: {
      readOnly: false, inputTime: 101, undoGroup: 'mode',
    } });
    const x = value.store.dispatch(insert(value.selection, 'Y'));
    const y = value.store.dispatch(insert(value.selection, 'Z'));
    expect(value.calls).toHaveLength(1);
    expect(value.store.getSnapshot().waitingInputCount).toBe(2);
    await value.waitForReplies(1);
    value.releaseOne();
    await first;
    await value.waitForReplies(2);
    expect(value.calls).toHaveLength(3);
    const sent = value.calls.map((call) => sharedWire(call.envelope));
    expect(sent.map((envelope) => envelope.seq)).toEqual([value.baseSeq + 1, value.baseSeq + 2, value.baseSeq + 3]);
    expect(sent[1]!.afterSeq).toBeUndefined();
    expect(sent[2]).toMatchObject({ afterSeq: sent[1]!.seq, payload: { text: 'YZ',
      at: { source: 'afterSeq', seq: sent[1]!.seq } } });
    expect(sent[2]!.basis).toEqual(sent[0]!.basis);
    value.release();
    const [mergedX, mergedY] = await Promise.all([x, y, independent]);
    expect(mergedX.outcomes).toEqual(mergedY.outcomes);
    expect(value.store.getSnapshot().activeLine?.text).toBe('aXYZbc');
  } finally { await value.dispose(); }
});

test('rec_disabled_recovery_never_reads_source_or_reopens', async () => {
  const value = await open();
  try {
    value.hold();
    const before = value.sourceReads;
    const frame = value.store.getSnapshot().confirmed;
    const operation = value.store.dispatch(insert(value.selection, 'X')).catch((error: unknown) => error);
    await value.waitForReplies(1);
    value.crash();
    const failure = await operation;
    expect(failure).toMatchObject({ name: 'SessionFailure', code: 'crash' });
    expect(value.workerFactory).toHaveBeenCalledTimes(1);
    expect(value.sourceReads).toBe(before);
    expect(value.errors).toHaveBeenCalledTimes(1);
    expect(value.methods).not.toContain('checkpoint');
    expect(value.store.getSnapshot().editState.readOnly).toBe(true);
    expect(value.store.getSnapshot().confirmed).toBe(frame);
    sharedWire(value.calls[0]!.envelope);
  } finally { await value.dispose(); }
});

test('rec_shared_path_holds_257th_input_without_losing_text', async () => {
  const value = await open({}, true);
  try {
    value.hold();
    const operations = Array.from({ length: 256 }, (_, index) =>
      value.store.dispatch(insert(value.selection, 'a', `typing-${index}`)));
    operations.push(value.store.dispatch(insert(value.selection, 'LAST')));
    expect(value.calls.length).toBe(256);
    expect(value.store.getSnapshot().waitingInputCount).toBe(1);
    expect(value.store.getSnapshot().echo?.text).toBe(`a${'a'.repeat(256)}bc`);
    await value.waitForReplies(256);
    value.releaseOne();
    await operations[0];
    await value.waitForReplies(256);
    expect(value.calls.length).toBe(257);
    const sent = sharedWire(value.calls[256]!.envelope);
    expect(sent).toMatchObject({ seq: value.baseSeq + 257, afterSeq: value.baseSeq + 256,
      payload: { text: 'LAST' } });
    expect(sent.basis).toEqual(sharedWire(value.calls[0]!.envelope).basis);
    value.release();
    await Promise.all(operations);
    expect(value.store.getSnapshot()).toMatchObject({ waitingInputCount: 0, pending: [],
      echo: null, presentedThroughSeq: value.baseSeq + 257, selection: { anchor: 261, focus: 261 } });
    const text = value.store.getSnapshot().confirmed!.displayList.primitives.flatMap((primitive) =>
      primitive.kind === 'textBox' && primitive.storyId === value.selection.storyId
        ? primitive.lines.flatMap((line) => line.runs.map((run) => run.text)) : []).join('');
    expect(text).toBe(`a${'a'.repeat(256)}LASTbc`);
  } finally { await value.dispose(); }
}, 30_000);

test('rec_rejected_install_preserves_coverage_echo_and_waiters', async () => {
  let reject = false;
  const value = await open({ paintFrame() { if (reject) throw new Error('Paint declined'); } }, true);
  try {
    value.hold();
    const operation = value.store.dispatch(insert(value.selection, 'a'));
    await value.waitForReplies(1);
    const before = value.store.getSnapshot();
    expect(before.echo?.text).toBe('aabc');
    const candidate = value.replyFrame();
    expect(candidate).toBeDefined();
    if (!candidate) throw new Error('Expected covering frame');
    let presented = false;
    const completion = deferred<void>();
    value.store.subscribePresentation(() => {
      if (value.store.getSnapshot().presentedThroughSeq >= value.baseSeq + 1) {
        presented = true;
        completion.resolve();
      }
    });
    reject = true;
    value.releaseOne();
    await operation;
    const refused = value.store.getSnapshot();
    expect(refused.appliedThroughSeq).toBe(value.baseSeq + 1);
    expect(refused.presentedThroughSeq).toBe(before.presentedThroughSeq);
    expect(refused.confirmed).toBe(before.confirmed);
    expect(refused.echo).toEqual(before.echo);
    expect(refused.pending.map((pending) => pending.seq)).toEqual(before.pending.map((pending) => pending.seq));
    await Promise.resolve();
    expect(presented).toBe(false);
    expect(value.accepted[value.accepted.length - 1]).toBe(before.confirmed!);
    reject = false;
    expect(installCoveringFrame(value.store, candidate)).toBe(true);
    value.release();
    await operation;
    await completion.promise;
    expect(value.store.getSnapshot().activeLine?.text).toBe('aabc');
    expect(value.store.getSnapshot().echo).toBeNull();
  } finally { await value.dispose(); }
});

test('rec_undecodable_covering_ack_rejects_dispatch_and_releases_fence', async () => {
  const errors = mock((_error: Error) => {});
  const value = await open({}, true, errors);
  try {
    value.hold();
    const before = value.store.getSnapshot().confirmed;
    let settlement = 'pending';
    let failure: unknown;
    const operation = value.store.dispatch(insert(value.selection, 'a')).then(
      () => { settlement = 'resolved'; },
      (error: unknown) => { settlement = 'rejected'; failure = error; },
    );
    let fenceSettlement = 'pending';
    let fenceFailure: unknown;
    void value.store.fence().then(
      () => { fenceSettlement = 'resolved'; },
      (error: unknown) => { fenceSettlement = 'rejected'; fenceFailure = error; },
    );
    await value.waitForReplies(1);
    const candidate = value.replyFrame();
    if (!candidate) throw new Error('Expected covering frame');
    candidate.displayList = new ArrayBuffer(0);
    value.releaseOne();
    for (let index = 0; index < 32; index += 1) await Promise.resolve();
    expect(settlement).toBe('rejected');
    expect(failure).toMatchObject({ name: 'PptxFrameDecodeError', seq: value.baseSeq + 1 });
    expect(fenceSettlement).toBe('rejected');
    expect(fenceFailure === failure).toBe(true);
    expect(errors.mock.calls.length).toBe(1);
    expect(errors.mock.calls[0][0] === failure).toBe(true);
    expect(value.store.getSnapshot().confirmed === before).toBe(true);
    let fenced = false;
    void value.store.fence().then(() => { fenced = true; });
    for (let index = 0; index < 8; index += 1) await Promise.resolve();
    expect(fenced).toBe(true);
    const later = value.store.dispatch(insert(value.selection, 'b', 'later'));
    await value.waitForReplies(1);
    value.release();
    await later;
    await operation;
    expect(value.store.getSnapshot().activeLine?.text).toBe('aabbc');
    expect(value.store.getSnapshot().pending.length).toBe(0);
    await value.store.fence();
    await value.client.flush();
    const zip = await JSZip.loadAsync(await value.client.save());
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slide.includes('aabbc')).toBe(true);
  } finally { await value.dispose(); }
});

test('rec_newer_same_coverage_active_frame_preserves_pending_echo', async () => {
  const value = await open({}, true);
  try {
    value.hold();
    const operation = value.store.dispatch(insert(value.selection, 'a'));
    await value.waitForReplies(1);
    const before = value.store.getSnapshot();
    expect(before.echo?.text).toBe('aabc');
    const frame = { ...before.confirmed!.frame, epoch: before.confirmed!.frame.epoch + 1,
      slideRevision: before.confirmed!.frame.slideRevision + 1 };
    expect(installCoveringFrame(value.store, frame)).toBe(true);
    const after = value.store.getSnapshot();
    expect(after.confirmed!.frame).toBe(frame);
    expect(after.presentedThroughSeq).toBe(before.presentedThroughSeq);
    expect(after.echo).toEqual(before.echo);
    expect(after.pending).toEqual(before.pending);
    expect(peerlessInputReference(value.store)).toEqual({ source: 'afterSeq', seq: value.baseSeq + 1 });
    sharedWire(value.calls[0]!.envelope);
    value.release();
    await operation;
    expect(value.store.getSnapshot().activeLine?.text).toBe('aabc');
  } finally { await value.dispose(); }
});

test('rec_thumbnail_with_same_coverage_never_retires_active_echo', async () => {
  const value = await open({}, true);
  try {
    const slideId = value.client.state.slides[1]!.id;
    const reply = await value.client.thumbnail(slideId);
    expect(reply.status).toBe('frame');
    if (reply.status !== 'frame') throw new Error('Expected thumbnail frame');
    const thumbnail = reply.frame;
    value.hold();
    const operation = value.store.dispatch(insert(value.selection, 'a'));
    await value.waitForReplies(1);
    const before = value.store.getSnapshot();
    expect(before.echo?.text).toBe('aabc');
    expect(thumbnail.coversThroughSeq).toBe(before.presentedThroughSeq);
    const count = value.accepted.length;
    expect(value.store.installThumbnail(thumbnail)).toBe(true);
    const after = value.store.getSnapshot();
    expect(after.thumbnails.get(slideId)?.frame).toBe(thumbnail);
    expect(after.confirmed).toBe(before.confirmed);
    expect(after.presentedThroughSeq).toBe(before.presentedThroughSeq);
    expect(after.echo).toEqual(before.echo);
    expect(after.pending).toEqual(before.pending);
    expect(value.accepted).toHaveLength(count);
    sharedWire(value.calls[0]!.envelope);
    value.release();
    await operation;
    expect(value.store.getSnapshot().activeLine?.text).toBe('aabc');
  } finally { await value.dispose(); }
});

import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { isCommandAck, isCommandEnvelope, type CommandEnvelope } from '../../../../shared/office-session/commands/types';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import type { PptxReadResult } from '../edits';
import type { SlideDisplayList, TextBoxPrimitive } from '../types';
import { initWasm } from '../wasm/loader';
import { createSharedPresentationCommandClient } from './commandClient';
import { createPresentationCommandProjection } from './commandSharedAdapter';
import { createPeerlessSession } from './commandSession';
import type {
  PresentationCommandAck, PresentationCommandFence, PresentationCommandFrame, PresentationCommandIntent,
  PresentationCommandOpenResult, PresentationCommandSelection,
} from './commandProtocol';
import { createPresentationSessionHost } from './host';

const input = { inputTime: 100, undoGroup: 'typing' };

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

function content(result: PptxReadResult) {
  expect(result.ok).toBe(true);
  if (!result.ok) throw new Error(result.failure.message);
  return result;
}

function insert(selection: PresentationCommandSelection, text: string): PresentationCommandIntent {
  return { kind: 'insertText', payload: { ...input, text, at: { source: 'confirmed', selection } } };
}

function selected(selection: PresentationCommandSelection): PresentationCommandIntent {
  return { kind: 'setSelection', payload: { ...input, selection } };
}

function sharedAck(value: unknown) {
  expect(isCommandAck(value)).toBe(true);
  if (!isCommandAck(value)) throw new Error('Expected shared acknowledgement');
  return value;
}

function sharedEnvelope(value: unknown) {
  expect(isCommandEnvelope(value)).toBe(true);
  if (!isCommandEnvelope(value)) throw new Error('Expected shared envelope');
  return value;
}

async function open() {
  const pair = createInProcessPair();
  let disposeHost = () => {};
  const hostTransport: SessionTransport = { ...pair.host, listen(listener) {
    disposeHost = () => listener({ protocol: 1, kind: 'dispose' });
    return pair.host.listen(listener);
  } };
  const calls: { id: number; method: string; envelope: unknown }[] = [];
  const replies = new Map<number, unknown>();
  const transport: SessionTransport = { ...pair.client,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call' &&
        (message.method === 'command' || message.method === 'edit')) {
        calls.push({ id: message.id, method: message.method, envelope: message.args[0] });
      }
      pair.client.post(message, transfer);
    },
    listen(listener) {
      return pair.client.listen((message) => {
        if (isHostMessage(message) && message.kind === 'reply') {
          replies.set(message.id, message.ok ? message.value : message.error);
        }
        listener(message);
      });
    },
  };
  createPresentationSessionHost(hostTransport, { initWasm: async () => {
    await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  const events = new EventTarget();
  const off = transport.listen((message) => events.dispatchEvent(new MessageEvent('message', { data: message })));
  const worker = {
    postMessage(message: unknown, transfer?: Transferable[]) { transport.post(message, transfer); },
    addEventListener: events.addEventListener.bind(events),
    removeEventListener: events.removeEventListener.bind(events),
    terminate() { disposeHost(); off(); transport.close(); },
  } as unknown as Worker;
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const owner = createPeerlessSession(await deck(), { worker: () => worker, clientId: 7811,
    fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }],
  });
  const client = await owner.ready;
  const rpc = createSessionClient<{ edit(envelope: CommandEnvelope<unknown>): unknown }, Record<string, never>>(
    pair.client, { methods: { edit: true } },
  );
  const story = content((await client.read()).result).stories[0]!;
  const selection: PresentationCommandSelection = { slideId: story.slideId, shapeId: story.shapeId,
    storyId: story.storyId, anchor: 1, focus: 1, affinity: 'downstream' };
  return { client, rpc, selection, calls,
    reply(index: number) { return replies.get(calls[index]!.id); },
    async dispose() { try { await owner.dispose(); } finally { await rpc.dispose(); } },
  };
}

test('rec_shared_connection_accepts_initial_and_edit_acks', async () => {
  const value = await open();
  try {
    const initial = value.client.state;
    expect(initial).toMatchObject({ throughSeq: 0, outcomes: [], dirtySlideIds: [] });
    await value.client.dispatch(selected(value.selection));
    const sent = value.calls[0]!.envelope;
    const probe: CommandEnvelope<unknown> = isCommandEnvelope(sent) ? sent : {
      session: initial.session, incarnation: 0, seq: 1, basis: { epoch: initial.epoch, version: 0 },
      ...selected(value.selection),
    };
    const baseline = sharedAck(await value.rpc.call.edit(probe).catch((error: unknown) => error));
    expect(baseline).toMatchObject({ throughSeq: 1, frame: { coversThroughSeq: 1 } });
    expect(baseline.frame?.payload).toMatchObject({ slideId: initial.activeSlideId, version: initial.version,
      coversThroughSeq: 1, documentEpoch: initial.epoch, incarnation: initial.incarnation });
    expect(baseline.frame?.revision).toBe((baseline.frame?.payload as PresentationCommandFrame).epoch);
    const ack = await value.client.dispatch(insert(value.selection, 'X'));
    const edited = sharedAck(value.reply(1));
    expect(edited).toMatchObject({ session: baseline.session, incarnation: baseline.incarnation,
      epoch: baseline.epoch, throughSeq: 2, version: baseline.version + 1 });
    expect(edited.frame?.payload).toEqual(ack.frame);
    expect(typeof ack.version).toBe('string');
    expect(content((await value.client.read()).result).stories[0]!.text).toBe('aXbc');
  } finally { await value.dispose(); }
});

test('rec_shared_projection_preserves_committed_and_dirty_slides', async () => {
  const value = await open();
  try {
    const ack = await value.client.dispatch(insert(value.selection, 'X'));
    expect(ack.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    expect(ack.dirtySlideIds).toEqual([value.selection.slideId]);
    const shared = sharedAck(value.reply(0));
    expect(shared.outcomes).toEqual([{ seq: 1, status: 'applied' }]);
    expect(shared.frame).toMatchObject({ coversThroughSeq: 1, revision: ack.frame!.epoch });
    const frame = shared.frame?.payload as PresentationCommandFrame;
    expect(frame).toEqual(ack.frame!);
    const list = JSON.parse(new TextDecoder().decode(frame.displayList)) as SlideDisplayList;
    expect(list.primitives.filter((primitive): primitive is TextBoxPrimitive => primitive.kind === 'textBox')
      .flatMap((box) => box.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join(''))))
      .toEqual(['aXbc']);
  } finally { await value.dispose(); }
});

test('rec_state_only_applied_preserves_document_clock', async () => {
  const value = await open();
  try {
    await value.client.dispatch(selected(value.selection));
    const baseline = sharedAck(value.reply(0));
    const version = value.client.state.version;
    const ack = await value.client.dispatch({ kind: 'setReadOnly', payload: { ...input, readOnly: true } });
    const shared = sharedAck(value.reply(1));
    expect(shared.version).toBe(baseline.version);
    expect(ack.version).toBe(version);
    expect(ack).toMatchObject({ throughSeq: 2, dirtySlideIds: [], editState: { readOnly: true },
      outcomes: [{ seq: 2, status: 'applied', committed: false }] });
    expect(shared.outcomes).toEqual([{ seq: 2, status: 'applied', committed: false }]);
    expect(content((await value.client.read()).result).stories[0]!.text).toBe('abc');
  } finally { await value.dispose(); }
});

test('rec_duplicate_and_pruned_duplicate_preserve_pptx_outcomes', async () => {
  const value = await open();
  try {
    const first = await value.client.dispatch(insert(value.selection, 'X'));
    const envelope = sharedEnvelope(value.calls[0]!.envelope);
    const original = sharedAck(value.reply(0));
    const retained = sharedAck(await value.rpc.call.edit(envelope).catch((error: unknown) => error));
    expect(retained.outcomes).toEqual(original.outcomes);
    expect(retained.version).toBe(original.version);
    expect(first.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    for (let index = 0; index < 257; index += 1) await value.client.dispatch(selected(value.selection));
    const latest = sharedAck(value.reply(value.calls.length - 1));
    const expired = sharedAck(await value.rpc.call.edit(envelope).catch((error: unknown) => error));
    expect(expired.outcomes).toEqual([{ seq: 1, status: 'refused', reason: 'stale-basis', committed: false }]);
    expect(expired.version).toBe(latest.version);
    expect(expired.throughSeq).toBe(latest.throughSeq);
    expect(content((await value.client.read()).result).stories[0]!.text).toBe('aXbc');
  } finally { await value.dispose(); }
}, 30_000);

test('rec_confirmed_reference_keeps_captured_basis', async () => {
  const value = await open();
  try {
    const basis = { epoch: value.client.state.epoch, version: value.client.state.version };
    await value.client.dispatch(insert(value.selection, 'X'), basis);
    const first = sharedEnvelope(value.calls[0]!.envelope);
    const ack = await value.client.dispatch(insert({ ...value.selection, anchor: 2, focus: 2 }, 'Y'), basis);
    const queued = sharedEnvelope(value.calls[1]!.envelope);
    expect(queued.basis).toEqual(first.basis);
    expect(queued.afterSeq).toBeUndefined();
    expect(queued.expectVersion ?? false).toBe(false);
    expect(queued.payload).toMatchObject({ at: { source: 'confirmed', selection: { anchor: 2, focus: 2 } } });
    expect(sharedAck(value.reply(1)).version).toBe(sharedAck(value.reply(0)).version + 1);
    expect(ack.outcomes).toEqual([{ seq: 2, status: 'applied', committed: true }]);
    expect(content((await value.client.read()).result).stories[0]!.text).toBe('aXbYc');
  } finally { await value.dispose(); }
});

test('rec_shared_fences_include_text_blocked_behind_256_replies', async () => {
  const selection: PresentationCommandSelection = { slideId: 'slide', shapeId: 'box', storyId: 'story',
    anchor: 1, focus: 1, affinity: 'downstream' };
  const state: PresentationCommandOpenResult = {
    session: 'session', incarnation: 'worker', epoch: 1, throughSeq: 0, version: 'v0', selection,
    editState: { readOnly: false, canUndo: false, canRedo: false, toolbar: { target: selection,
      storedMarks: {}, formatting: null, alignment: null, hyperlink: null, table: null, canEditText: true } },
    outcomes: [], dirtySlideIds: [], activeSlideId: 'slide', slides: [], size: { width: 960, height: 540 },
    frame: { slideId: 'slide', slideIndex: 0, slideRevision: 0, coversThroughSeq: 0, documentEpoch: 1,
      incarnation: 'worker', epoch: 1, version: 'v0', displayList: new ArrayBuffer(0), media: [] },
  };
  const projection = createPresentationCommandProjection();
  projection.initialize(state.version);
  let deliver: (message: unknown) => void = () => {};
  let ack: PresentationCommandAck = state;
  let text = 'abc';
  const texts = new Map([[0, text]]);
  const acks = new Map<number, PresentationCommandAck>([[0, state]]);
  const held: (() => void)[] = [];
  const fences: number[] = [];
  const transport: SessionTransport = {
    post(message) {
      if (!isClientMessage(message) || message.kind !== 'call') return;
      const reply = (value: unknown) => deliver({ protocol: 1, kind: 'reply', id: message.id, ok: true, value });
      if (message.method === 'openCommandEditor') reply({ ...state, shared: projection.ack(state) });
      else if (message.method === 'edit') {
        const command = sharedEnvelope(message.args[0]);
        const payload = command.payload as Extract<PresentationCommandIntent, { kind: 'insertText' }>['payload'];
        const position = payload.at.source === 'confirmed' ? payload.at.selection.focus : ack.selection!.focus;
        text = text.slice(0, position) + payload.text + text.slice(position);
        projection.commit(`v${command.seq}`);
        ack = { ...ack, throughSeq: command.seq, version: `v${command.seq}`,
          selection: { ...selection, anchor: position + payload.text.length, focus: position + payload.text.length },
          outcomes: [{ seq: command.seq, status: 'applied', committed: true }] };
        texts.set(command.seq, text);
        acks.set(command.seq, ack);
        const value = projection.ack(ack);
        if (command.seq <= 256) held.push(() => reply(value));
        else reply(value);
      } else {
        const fence = message.args[0] as PresentationCommandFence;
        fences.push(fence.seq);
        const value = acks.get(fence.seq)!;
        if (message.method === 'saveCommandEditor') reply({ ...value,
          bytes: new TextEncoder().encode(texts.get(fence.seq)!).buffer });
        else if (message.method === 'readEditor') reply({ ...value,
          result: { ok: true, stories: [{ text: texts.get(fence.seq) }] } });
        else reply(value);
      }
    },
    listen(listener) { deliver = listener; return () => { deliver = () => {}; }; },
    onError() { return () => {}; },
    close() { held.length = 0; },
  };
  const client = await createSharedPresentationCommandClient(new Uint8Array(),
    { session: state.session, incarnation: state.incarnation, clientId: 1 }, transport);
  try {
    const operations = Array.from({ length: 256 }, () => client.dispatch(insert(selection, 'x')));
    operations.push(client.dispatch(insert(selection, 'LAST')));
    expect(held).toHaveLength(256);
    expect(text).not.toContain('LAST');
    let saved: Uint8Array | undefined;
    let flushed: PresentationCommandAck | undefined;
    let readText: string | undefined;
    const barriers = [client.save().then((value) => { saved = value; }),
      client.flush().then((value) => { flushed = value; }),
      client.read().then((value) => { readText = content(value.result).stories[0]!.text; })];
    let completed = false;
    void Promise.all([...operations, ...barriers]).then(() => { completed = true; }, () => {});
    for (const reply of held.splice(0)) reply();
    for (let turn = 0; turn < 32; turn += 1) await Promise.resolve();
    expect(new TextDecoder().decode(saved)).toBe(`aLAST${'x'.repeat(256)}bc`);
    expect(flushed?.throughSeq).toBe(257);
    expect(readText).toBe(`aLAST${'x'.repeat(256)}bc`);
    expect(fences).toEqual([257, 257, 257]);
    expect(completed).toBe(true);
  } finally { await client.dispose(); projection.dispose(); }
});

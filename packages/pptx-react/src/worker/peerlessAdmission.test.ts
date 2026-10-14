import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionHost } from '../../../../shared/office-session/host';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { createPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import type {
  PresentationCommandAck, PresentationCommandEvents, PresentationCommandIntent,
  PresentationCommandMethods, PresentationCommandSelection, PresentationCommandMotion,
} from '@betteroffice/pptx';
import { initWasm } from '../../../pptx/src/wasm/loader';
import { commitComposition, createPeerlessViewStore, navigateConfirmedSelection,
  PEERLESS_PENDING_INTENTS, peerlessInputReference } from './peerlessViewStore';

async function deck(text: string): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="8000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="none"/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function session(text = 'abc', position = 1, options: { echoEnabled?: boolean } = {}) {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
    handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
  });
  const commandIds = new Set<number>();
  const held: unknown[] = [];
  let deliver: ((message: unknown) => void) | undefined;
  let holding = false;
  let expected = 0;
  let reached: (() => void) | undefined;
  const transport: SessionTransport = { ...pair.client,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call' && message.method === 'command') commandIds.add(message.id);
      pair.client.post(message, transfer);
    },
    listen(listener) {
      deliver = listener;
      return pair.client.listen((message) => {
        if (holding && isHostMessage(message) && message.kind === 'reply' && commandIds.has(message.id)) {
          held.push(message);
          if (held.length >= expected) reached?.();
        } else listener(message);
      });
    },
  };
  const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
  const client = await createPresentationCommandClient(await deck(text), {
    session: 'p2-session', incarnation: 'p2-worker', clientId: 7822,
    fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }],
  }, transport);
  const content = (await client.read()).result;
  if (!content.ok || !content.stories[0]) throw new Error('Expected text story');
  const story = content.stories[0];
  const at: PresentationCommandSelection = { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
    anchor: position, focus: position, affinity: 'downstream' };
  await client.dispatch({ kind: 'replaceText', payload: {
    at: { source: 'confirmed', selection: { ...at, anchor: 0, focus: text.length } }, text,
    style: { bold: false, italic: false, underline: 'none', fontSizePt: 12,
      color: '#000000', fontFamily: 'Liberation Sans' }, inputTime: 0, undoGroup: 'setup',
  } });
  await client.dispatch({ kind: 'setSelection', payload: { selection: at, inputTime: 0, undoGroup: 'selection' } });
  const store = createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} }, options);
  return { client, store, at,
    hold(count: number) {
      holding = true;
      expected = count;
      return new Promise<void>((resolve) => { reached = resolve; });
    },
    release() {
      holding = false;
      const latest = held.pop();
      if (latest) deliver?.(latest);
      for (const reply of held.splice(0)) deliver?.(reply);
    },
    async savedText() {
      await store.fence();
      const zip = await JSZip.loadAsync(await client.save());
      const slide = zip.file('ppt/slides/slide1.xml');
      if (!slide) throw new Error('Expected saved slide');
      return [...(await slide.async('string')).matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)].map((match) => match[1]).join('');
    },
    async dispose() { store.dispose(); try { await client.dispose(); } finally { registration.dispose(); pair.client.close(); } },
  };
}

function type(store: ReturnType<typeof createPeerlessViewStore>, text: string): Promise<PresentationCommandAck> {
  const at = peerlessInputReference(store);
  if (!at) throw new Error('Expected input reference');
  return store.dispatch({ kind: 'insertText', payload: { at, text, inputTime: 100, undoGroup: 'typing' } });
}

test('p2_real_host_admits_and_saves_input_queued_behind_acknowledged_cap', async () => {
  const value = await session('abc', 1, { echoEnabled: true });
  try {
    const held = value.hold(PEERLESS_PENDING_INTENTS);
    const pending = Array.from({ length: PEERLESS_PENDING_INTENTS }, () => type(value.store, 'a'));
    const queued = type(value.store, 'LAST');
    expect(value.store.getSnapshot().waitingInputCount).toBe(1);
    await held;
    value.release();
    const outcomes = await Promise.all([...pending, queued]);
    const last = outcomes[outcomes.length - 1];
    expect(last?.outcomes[last.outcomes.length - 1]).toMatchObject({ status: 'applied', committed: true });
    expect(await value.savedText()).toBe(`a${'a'.repeat(256)}LASTbc`);
  } finally { await value.dispose(); }
});

test('p2_real_host_coalesces_and_saves_typing_behind_unacknowledged_text', async () => {
  const value = await session();
  try {
    const held = value.hold(1);
    const pending = Array.from({ length: PEERLESS_PENDING_INTENTS }, () => type(value.store, 'a'));
    const queued = type(value.store, 'LAST');
    expect(value.store.getSnapshot().echo).toBeNull();
    expect(value.store.getSnapshot().pending).toHaveLength(1);
    expect(value.store.getSnapshot().pending[0]).toMatchObject({ status: 'pending', echoed: false });
    expect(value.store.getSnapshot().waitingInputCount).toBe(1);
    await held;
    value.release();
    const outcomes = await Promise.all([...pending, queued]);
    expectApplied(outcomes, PEERLESS_PENDING_INTENTS + 1);
    const last = outcomes[outcomes.length - 1];
    expect(last?.outcomes[last.outcomes.length - 1]).toMatchObject({ status: 'applied', committed: true });
    expect(await value.savedText()).toBe(`a${'a'.repeat(256)}LASTbc`);
    expect(value.store.getSnapshot()).toMatchObject({ pending: [], waitingInputCount: 0, echo: null });
  } finally { await value.dispose(); }
});

test('p2_real_host_admits_and_saves_input_queued_behind_oversized_paste', async () => {
  const value = await session();
  try {
    const held = value.hold(1);
    const paste = 'P'.repeat(8193);
    const at = peerlessInputReference(value.store);
    if (!at) throw new Error('Expected paste reference');
    const intent: PresentationCommandIntent = { kind: 'replaceText', payload: {
      at, text: paste, inputTime: 100, undoGroup: 'paste',
    } };
    const pasted = value.store.dispatch(intent);
    const queued = type(value.store, 'LAST');
    expect(value.store.getSnapshot().waitingInputCount).toBe(1);
    await held;
    value.release();
    const [, ack] = await Promise.all([pasted, queued]);
    expect(ack.outcomes[ack.outcomes.length - 1]).toMatchObject({ status: 'applied', committed: true });
    expect(await value.savedText()).toBe(`a${paste}LASTbc`);
  } finally { await value.dispose(); }
});

test('p2_real_host_types_at_explicit_caret_while_text_is_pending', async () => {
  const value = await session();
  try {
    const held = value.hold(1);
    const inserted = type(value.store, 'a');
    const caret = click(value.store, 3);
    const typed = type(value.store, '!');
    expect(value.store.getSnapshot().echo).toBeNull();
    expect(value.store.getSnapshot().pending).toHaveLength(1);
    expect(value.store.getSnapshot().pending[0]).toMatchObject({ status: 'pending', echoed: false });
    expect(value.store.getSnapshot().waitingInputCount).toBe(2);
    await held;
    value.release();
    expectApplied(await Promise.all([inserted, caret, typed]), 3);
    expect(await value.savedText()).toBe('aabc!');
    expect(value.store.getSnapshot().selection).toMatchObject({ anchor: 5, focus: 5 });
    expect(value.store.getSnapshot().activeLine?.text).toBe('aabc!');
    expect(value.store.getSnapshot()).toMatchObject({ pending: [], waitingInputCount: 0, echo: null });
  } finally { await value.dispose(); }
});

function move(store: ReturnType<typeof createPeerlessViewStore>, motion: PresentationCommandMotion, extend = false) {
  const at = peerlessInputReference(store);
  if (!at) throw new Error('Expected selection reference');
  return store.dispatch({ kind: 'moveSelection', payload: { at, motion, extend, inputTime: 100, undoGroup: 'selection' } });
}

function click(store: ReturnType<typeof createPeerlessViewStore>, position: number) {
  const at = peerlessInputReference(store);
  const confirmed = store.getSnapshot().confirmed!;
  const box = confirmed.displayList.primitives.find((primitive) => primitive.kind === 'textBox');
  if (!at || box?.kind !== 'textBox') throw new Error('Expected frame text geometry');
  const line = box.lines.find((line) => line.caretStops.some((stop) => stop.position === position));
  const stop = line?.caretStops.find((stop) => stop.position === position);
  if (!line || !stop) throw new Error('Expected caret stop');
  const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = confirmed.frame;
  return store.dispatch({ kind: 'selectAtPoint', payload: { at,
    frame: { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation },
    x: stop.x, y: line.y + line.height / 2, extend: false, inputTime: 100, undoGroup: 'pointer' } });
}

function expectApplied(acks: PresentationCommandAck[], count: number) {
  expect(acks).toHaveLength(count);
  for (const ack of acks) expect(ack.outcomes[ack.outcomes.length - 1]?.status).toBe('applied');
}

test('p3_real_host_queues_echo_end_and_following_text_until_its_basis_is_acknowledged', async () => {
  const value = await session('abc', 3);
  try {
    const held = value.hold(1);
    const inserted = type(value.store, 'a');
    const end = navigateConfirmedSelection(value.store, 'End');
    expect(end).toBeNull();
    const selected = move(value.store, 'lineEnd');
    const typed = type(value.store, 'Y');
    await held;
    value.release();
    expectApplied(await Promise.all([inserted, selected, typed]), 3);
    expect(await value.savedText()).toBe('abcaY');
    expect(value.store.getSnapshot()).toMatchObject({ pending: [], waitingInputCount: 0 });
  } finally { await value.dispose(); }
});

test('p3_real_host_rebases_confirmed_input_after_a_deletion_without_echo', async () => {
  const text = `abcdef${'z'.repeat(8193)}`;
  const value = await session(text);
  try {
    expect(value.store.getSnapshot().activeLine).toBeNull();
    const held = value.hold(1);
    const deleted = value.store.dispatch({ kind: 'deleteForward', payload: {
      at: { source: 'confirmed', selection: value.at }, unit: 'codePoint', inputTime: 100, undoGroup: 'typing',
    } });
    const selected = click(value.store, 3);
    const typed = type(value.store, '!');
    await held;
    value.release();
    expectApplied(await Promise.all([deleted, selected, typed]), 3);
    expect(await value.savedText()).toBe(`ac!def${'z'.repeat(8193)}`);
    expect(value.store.getSnapshot()).toMatchObject({ pending: [], waitingInputCount: 0 });
  } finally { await value.dispose(); }
});

test('p3_real_host_keeps_oversized_paste_outside_a_range_ending_at_its_insertion', async () => {
  const value = await session();
  try {
    const held = value.hold(1);
    const paste = 'X'.repeat(8193);
    const pasted = value.store.dispatch({ kind: 'replaceText', payload: {
      at: { source: 'confirmed', selection: value.at }, text: paste, inputTime: 100, undoGroup: 'paste',
    } });
    const home = navigateConfirmedSelection(value.store, 'Home', true);
    expect(home).toBeNull();
    const selected = move(value.store, 'lineStart', true);
    const at = peerlessInputReference(value.store);
    if (!at) throw new Error('Expected replacement reference');
    const typed = value.store.dispatch({ kind: 'replaceText', payload: {
      at, text: 'Y', inputTime: 100, undoGroup: 'typing',
    } });
    await held;
    value.release();
    expectApplied(await Promise.all([pasted, selected, typed]), 3);
    expect(await value.savedText()).toBe('Ybc');
    expect(value.store.getSnapshot()).toMatchObject({ pending: [], waitingInputCount: 0 });
  } finally { await value.dispose(); }
});

test('p3_composition_selection_and_reference_rebase_together_before_preedit_rebuild', async () => {
  const value = await session();
  try {
    const held = value.hold(1);
    const inserted = type(value.store, 'a');
    const selected = click(value.store, 3);
    value.store.beginComposition('p3-ime', 100);
    value.store.updateComposition('c');
    expect(value.store.getSnapshot().echo).toBeNull();
    expect(value.store.getSnapshot().pending).toHaveLength(1);
    expect(value.store.getSnapshot().pending[0]).toMatchObject({ status: 'pending', echoed: false });
    expect(value.store.getSnapshot().waitingInputCount).toBe(2);
    await held;
    value.release();
    await value.store.fence();
    expectApplied(await Promise.all([inserted, selected]), 2);
    expect(value.store.getSnapshot().activeLine?.text).toBe('aabc');
    expect(value.store.getSnapshot().selection).toMatchObject({ anchor: 4, focus: 4 });
    expect(value.store.getSnapshot().composition?.selection).toMatchObject({ anchor: 4, focus: 4 });
    expect(value.store.getSnapshot().composition?.text).toBe('c');
    expect(value.store.getSnapshot()).toMatchObject({ pending: [], waitingInputCount: 0, echo: null });
    const committed = await commitComposition(value.store, 'c', 101);
    expect(committed?.outcomes[committed.outcomes.length - 1]?.status).toBe('applied');
    expect(await value.savedText()).toBe('aabcc');
    expect(value.store.getSnapshot().activeLine?.text).toBe('aabcc');
    expect(value.store.getSnapshot().selection).toMatchObject({ anchor: 5, focus: 5 });
    expect(value.store.getSnapshot()).toMatchObject({ composition: null, pending: [], waitingInputCount: 0, echo: null });
  } finally { await value.dispose(); }
});

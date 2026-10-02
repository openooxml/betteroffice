import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import {
  createYrsSession,
  type DocxEditReceipt,
  type DocxEditResult,
  type DocxEditStep,
  type YrsSession,
} from './index';
import { proposalRevisionPreview, type DocxProposalInput } from './proposals';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import { ResidentEngineWorkerClient, type ResidentProposalReply } from './residentEngineWorkerClient';
import type { ResidentProposalOperation } from './residentEngineWorkerProtocol';
import { saveYrsDocx } from './saveYrsDocx';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = new Uint8Array(readFileSync(resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});
const TEXTS = ['Alpha beta gamma', 'Delta epsilon', 'User', 'Remote'];
const SUGGEST = { author: 'Host', date: '2026-09-30T00:00:00Z' };
const sessions: YrsSession[] = [];
const clients: ResidentEngineWorkerClient[] = [];
let startWorker: (clientId?: number) => InProcessResidentWorker;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  startWorker = await residentWorkerFactory();
});

afterAll(() => {
  for (const client of clients.splice(0)) client.destroy();
  for (const session of sessions.splice(0)) session.destroy();
});

function proposalDocument(texts = TEXTS): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  ));
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>' +
    texts.map((text, index) =>
      `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}"><w:r><w:t>${text}</w:t></w:r></w:p>`
    ).join('') + '<w:sectPr/></w:body></w:document>'
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

interface Arm {
  bytes: Uint8Array;
  main: YrsSession;
  engine: ResidentEngineSession;
  client: ResidentEngineWorkerClient;
  frame: RetainedFrame;
  layoutInput: string;
  adopting: boolean;
  origins: string[];
}

async function replica(bytes: Uint8Array, state: Uint8Array, clientId: number): Promise<YrsSession> {
  const session = await createYrsSession({ clientId });
  sessions.push(session);
  session.openDocx(bytes, false);
  session.loadState(state);
  return session;
}

async function openArm(bytes: Uint8Array, clientId: number, direct: boolean): Promise<Arm> {
  const worker = startWorker(clientId);
  const client = new ResidentEngineWorkerClient(worker);
  clients.push(client);
  await client.open(bytes, { generation: 'direct-batches' });
  expect(worker.sessions).toHaveLength(1);
  const engine = worker.sessions[0]!;
  if (!direct) engine.setDirectBatches(false);
  expect(engine.directBatchesApplied()).toBe(0);
  const main = await replica(bytes, await client.encodeState(), clientId + 1000);
  main.registerFont(FONT);
  main.adoptResidentWorkerLayout!(LAYOUT);
  main.beginUndoCapture();
  const booted = await client.bootstrap(
    { ...main.residentWorkerSnapshot()!, workerAuthoritative: true },
    '',
    { opened: true, layoutExtras: '{}' }
  );
  const arm: Arm = {
    bytes, main, engine, client,
    frame: applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame)),
    layoutInput: LAYOUT,
    adopting: false,
    origins: [],
  };
  main.onUpdate((update, origin) => {
    arm.origins.push(origin);
    if (!arm.adopting) client.invalidate(update, null);
  });
  await operate(arm, { kind: 'snapshot' });
  return arm;
}

async function comparison(bytes = proposalDocument()): Promise<[Arm, Arm]> {
  const direct = await openArm(bytes, 6101, true);
  const reference = await openArm(bytes, 6101, false);
  expectCounters(direct, reference, 0);
  await expectEquivalent(direct, reference);
  return [direct, reference];
}

function expectCounters(direct: Arm, reference: Arm, applied: number): void {
  expect(direct.engine.directBatchesApplied()).toBe(applied);
  expect(reference.engine.directBatchesApplied()).toBe(0);
}

function texts(session: YrsSession, view: 'accepted' | 'original'): string[] {
  const read = session.readParagraphs({ view });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map(({ text }) => text);
}

function replacement(id: string, paraId: string, search: string, replaceWith: string): DocxProposalInput {
  return {
    id,
    paragraph: {
      kind: 'persisted',
      story: { kind: 'body', partUri: '/word/document.xml' },
      paraId,
    },
    suggest: SUGGEST,
    op: 'replaceText', search, replaceWith,
  };
}

function batch(): DocxProposalInput[] {
  return [
    replacement('beta', '00000001', 'beta', 'BETA'),
    replacement('epsilon', '00000002', 'epsilon', 'EPSILON'),
  ];
}

async function operate(arm: Arm, operation: ResidentProposalOperation): Promise<ResidentProposalReply> {
  const reply = await arm.client.proposal(operation);
  expect(reply.mirror.version).toBe(arm.engine.proposalEngine.version());
  arm.adopting = true;
  try {
    for (const update of reply.updates) arm.main.applyHostUpdate(update, reply.changedStories);
    arm.main.mirrorWorkerDocument(reply.mirror);
  } finally {
    arm.adopting = false;
  }
  expect(arm.main.encodeStateVector()).toEqual(reply.stateVector);
  if (reply.result?.ok) {
    expect(reply.result.snapshot).toEqual(arm.main.getProposals());
  } else if (reply.result) {
    expect(reply.result.version).toBe(reply.mirror.version);
  }
  return reply;
}

async function propose(arm: Arm, proposals = batch()): Promise<ResidentProposalReply> {
  const reply = await operate(arm, {
    kind: 'propose',
    request: { expectVersion: arm.engine.proposalEngine.version(), proposals },
  });
  expect(reply.result?.ok).toBe(true);
  return reply;
}

async function expectWorkerParagraphs(arm: Arm, view: 'accepted' | 'original'): Promise<void> {
  const read = await arm.client.documentRead({ kind: 'readParagraphs', request: { view } });
  const main = arm.main.readParagraphs({ view });
  if (!read.value.ok || !main.ok) throw new Error('expected paragraph reads to succeed');
  expect(read.version).toBe(arm.engine.proposalEngine.version());
  expect(read.value.version).toBe(read.version);
  expect(read.value.view).toBe(main.view);
  expect(read.value.paragraphs).toEqual(main.paragraphs);
}

async function expectEquivalent(direct: Arm, reference: Arm): Promise<void> {
  expect(direct.main.encodeStateVector()).toEqual(reference.main.encodeStateVector());
  expect(direct.engine.encodeStateVector()).toEqual(direct.main.encodeStateVector());
  expect(reference.engine.encodeStateVector()).toEqual(reference.main.encodeStateVector());
  expect(direct.main.storySegments('body')).toEqual(reference.main.storySegments('body'));
  expect(direct.main.listRevisions()).toEqual(reference.main.listRevisions());
  expect(direct.main.getProposals().proposals).toEqual(reference.main.getProposals().proposals);
  expect(direct.main.getProposals().previewVersion).toBe(reference.main.getProposals().previewVersion);
  for (const view of ['accepted', 'original'] as const) {
    expect(texts(direct.main, view)).toEqual(texts(reference.main, view));
    for (const arm of [direct, reference]) await expectWorkerParagraphs(arm, view);
  }
  const identities = direct.main.paragraphIdentities();
  expect(reference.main.paragraphIdentities()).toEqual(identities);
  for (const arm of [direct, reference]) {
    const read = await arm.client.documentRead({ kind: 'paragraphIdentities' });
    expect(read.value).toEqual(identities);
    expect(arm.engine.geometryReader.storySegments('body')).toEqual(arm.main.storySegments('body'));
    expect(arm.engine.geometryReader.listRevisions()).toEqual(arm.main.listRevisions());
  }
}

async function relayout(arm: Arm): Promise<unknown> {
  const preview = proposalRevisionPreview(arm.main.getProposals());
  arm.layoutInput = JSON.stringify({
    ...JSON.parse(LAYOUT),
    renderEnv: preview ? { revisionPreview: preview } : {},
  });
  arm.main.adoptResidentWorkerLayout!(arm.layoutInput);
  const synced = await arm.client.sync(
    arm.main.residentWorkerSnapshot({
      knownStateVector: arm.client.remoteStateVector(),
      knownFontsRevision: arm.client.syncedFontsRevision(),
    })!,
    '', arm.frame.frameEpoch, false,
    { layoutExtras: '{}' }
  );
  arm.frame = applyFrameDeltaOwned(arm.frame, decodeFrameDelta(synced.frame));
  expect(synced.layoutJson).toBeDefined();
  return JSON.parse(synced.layoutJson!);
}

async function expectSaves(direct: Arm, reference: Arm): Promise<Uint8Array> {
  const left = await saveYrsDocx(direct.main);
  const right = await saveYrsDocx(reference.main);
  expect(left.conflicts).toEqual([]);
  expect(right.conflicts).toEqual([]);
  expect(right.bytes).toEqual(left.bytes);
  return left.bytes;
}

function expectHistory(direct: Arm, reference: Arm, undo: boolean, redo: boolean): void {
  for (const arm of [direct, reference]) {
    expect(arm.main.canUndo()).toBe(undo);
    expect(arm.main.canRedo()).toBe(redo);
  }
}

test('worker-open proposals stay outside Ctrl+Z history through decisions and withdrawal', async () => {
  const [direct, reference] = await comparison();
  expectHistory(direct, reference, false, false);
  const replies = [await propose(direct), await propose(reference)];
  expect(replies[0]!.changedStories).toEqual(['body']);
  expect(replies[1]!.changedStories).toEqual(replies[0]!.changedStories);
  expectCounters(direct, reference, 1);
  expectHistory(direct, reference, false, false);
  await expectEquivalent(direct, reference);
  const revisions = direct.main.listRevisions();
  const proposals = direct.main.getProposals().proposals;
  for (const arm of [direct, reference]) {
    expect(arm.main.undo()).toBe(false);
    await relayout(arm);
    const loc = { story: 'body', paraId: '00000003', offset: 4 };
    arm.main.setSelection(loc);
    const typed = await arm.client.applyInput(' word', arm.main.selection()!, arm.frame.frameEpoch);
    if (!typed.applied) throw new Error('the worker refused resident input');
    expect(typed.updates.length).toBeGreaterThan(0);
    arm.adopting = true;
    try {
      for (const update of typed.updates) arm.main.applyLocalUpdate(update);
    } finally {
      arm.adopting = false;
    }
    arm.frame = applyFrameDeltaOwned(arm.frame, decodeFrameDelta(typed.frame));
    if (typed.selection) arm.main.setSelection(typed.selection.anchor, typed.selection.head);
    await operate(arm, { kind: 'snapshot' });
    expect(texts(arm.main, 'accepted')).toEqual(['Alpha BETA gamma', 'Delta EPSILON', 'User word', 'Remote']);
    expect(arm.origins).toContain('local');
  }
  expectHistory(direct, reference, true, false);
  await expectEquivalent(direct, reference);
  for (const arm of [direct, reference]) {
    arm.main.addUndoBoundary();
    expect(arm.main.undo()).toBe(true);
    await relayout(arm);
    await operate(arm, { kind: 'snapshot' });
    expect(texts(arm.main, 'accepted')).toEqual(['Alpha BETA gamma', 'Delta EPSILON', 'User', 'Remote']);
    expect(arm.main.listRevisions()).toEqual(revisions);
    expect(arm.main.getProposals().proposals).toEqual(proposals);
    expect(arm.main.undo()).toBe(false);
  }
  expectHistory(direct, reference, false, true);
  await expectEquivalent(direct, reference);
  for (const arm of [direct, reference]) {
    expect(arm.main.redo()).toBe(true);
    await relayout(arm);
    await operate(arm, { kind: 'snapshot' });
    expect(texts(arm.main, 'accepted')).toEqual(['Alpha BETA gamma', 'Delta EPSILON', 'User word', 'Remote']);
    expect(arm.main.redo()).toBe(false);
    expect(arm.main.listRevisions()).toEqual(revisions);
  }
  expectHistory(direct, reference, true, false);
  await expectEquivalent(direct, reference);
  for (const [index, state] of (['accepted', 'proposed', 'rejected', 'proposed'] as const).entries()) {
    const decisions: ResidentProposalReply[] = [];
    for (const arm of [direct, reference]) {
      const reply = await operate(arm, { kind: 'setStates', request: {
        expectVersion: arm.engine.proposalEngine.version(),
        expectPreviewVersion: arm.main.getProposals().previewVersion,
        changes: [{ id: 'beta', state }],
      } });
      expect(reply.result?.ok).toBe(true);
      expect(reply.changedStories).toEqual([]);
      expect(reply.updates).toEqual([]);
      expect(arm.main.getProposals().previewVersion).toBe(index + 1);
      expect(arm.main.getProposals().proposals.find(({ id }) => id === 'beta')?.state).toBe(state);
      expect(arm.main.listRevisions()).toEqual(revisions);
      expect(texts(arm.main, 'accepted')).toEqual(['Alpha BETA gamma', 'Delta EPSILON', 'User word', 'Remote']);
      expect(texts(arm.main, 'original')).toEqual(['Alpha beta gamma', 'Delta epsilon', 'User word', 'Remote']);
      decisions.push(reply);
    }
    expect(decisions[0]!.mirror.proposals).toEqual(decisions[1]!.mirror.proposals);
    const layouts = [await relayout(direct), await relayout(reference)];
    expect(layouts[0]).toEqual(layouts[1]);
    expect(direct.frame.displayList).toEqual(reference.frame.displayList);
    expectCounters(direct, reference, 1);
    expectHistory(direct, reference, true, false);
    await expectEquivalent(direct, reference);
  }
  const withdrawn: ResidentProposalReply[] = [];
  for (const arm of [direct, reference]) {
    const reply = await operate(arm, { kind: 'withdraw', request: {
      expectVersion: arm.engine.proposalEngine.version(), ids: ['beta'],
    } });
    expect(reply.result?.ok).toBe(true);
    expect(reply.changedStories).toEqual(['body']);
    expect(texts(arm.main, 'accepted')).toEqual(['Alpha beta gamma', 'Delta EPSILON', 'User word', 'Remote']);
    expect(arm.main.getProposals().proposals.map(({ id }) => id)).toEqual(['epsilon']);
    withdrawn.push(reply);
  }
  expect(withdrawn[0]!.mirror.proposals).toEqual(withdrawn[1]!.mirror.proposals);
  expectCounters(direct, reference, 1);
  expectHistory(direct, reference, true, false);
  await expectEquivalent(direct, reference);
  await expectSaves(direct, reference);
});

test('direct batches emit the same story changes, identities and layout as replica batches and cold layout', async () => {
  const bytes = proposalDocument(Array.from({ length: 120 }, (_, index) =>
    `Paragraph ${index + 1} alpha beta gamma ${'delta epsilon '.repeat(16)}`
  ));
  const [direct, reference] = await comparison(bytes);
  const identities = direct.main.paragraphIdentities();
  const since = [direct, reference].map((arm) => ({
    main: arm.main.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision,
    worker: arm.engine.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision,
  }));
  const proposals = [1, 60, 120].map((index) => replacement(
    `p${index}`, index.toString(16).padStart(8, '0'), 'beta',
    'a much longer replacement that makes the paragraph wrap across more lines'
  ));
  const replies = [await propose(direct, proposals), await propose(reference, proposals)];
  expectCounters(direct, reference, 1);
  expect(replies[0]!.changedStories).toEqual(['body']);
  expect(replies[1]!.changedStories).toEqual(replies[0]!.changedStories);
  for (const reply of replies) expect(reply.updates.length).toBeGreaterThan(0);
  for (const [index, arm] of [direct, reference].entries()) {
    expect(arm.main.storiesChangedSince(since[index]!.main).stories).toEqual(['body']);
    expect(arm.engine.storiesChangedSince(since[index]!.worker).stories).toEqual(['body']);
    expect(arm.main.paragraphIdentities()).toEqual(identities);
  }
  await expectEquivalent(direct, reference);
  const layouts = [await relayout(direct), await relayout(reference)];
  expect(layouts[0]).toEqual(layouts[1]);
  expect(direct.frame.displayList.pages.length).toBeGreaterThan(2);
  expect(direct.frame.displayList).toEqual(reference.frame.displayList);
  const cold = await createResidentEngineSession(undefined, 6199);
  try {
    cold.loadState(await direct.client.encodeState());
    cold.registerFont(FONT);
    expect(JSON.parse(cold.layoutDocumentWithRegionsRetainedJson(direct.layoutInput))).toEqual(layouts[0]);
    const headersFooters = cold.retainedHeadersFootersJson();
    const extras = JSON.stringify(headersFooters ? { headersFooters: JSON.parse(headersFooters) } : {});
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(cold.buildDisplayListFrame(extras, 0)));
    expect(frame.displayList).toEqual(direct.frame.displayList);
    expect(cold.directBatchesApplied()).toBe(0);
  } finally {
    cold.destroy();
  }
  expectCounters(direct, reference, 1);
});

test('main replicas save identical bytes after a batch and a remote edit from the pre-batch state', async () => {
  const [direct, reference] = await comparison();
  const peer = await replica(direct.bytes, await direct.client.encodeState(), 6201);
  const before = peer.encodeStateVector();
  await propose(direct);
  await propose(reference);
  expectCounters(direct, reference, 1);
  await expectEquivalent(direct, reference);
  await expectSaves(direct, reference);
  peer.insertText({ story: 'body', paraId: '00000004', offset: 6 }, ' peer');
  const update = peer.encodeStateAsUpdate(before);
  for (const arm of [direct, reference]) {
    arm.main.applyUpdate(update);
    await relayout(arm);
    await operate(arm, { kind: 'snapshot' });
    expect(texts(arm.main, 'accepted')).toEqual(['Alpha BETA gamma', 'Delta EPSILON', 'User', 'Remote peer']);
    expect(arm.origins).toContain('remote');
  }
  expectCounters(direct, reference, 1);
  await expectEquivalent(direct, reference);
  await expectSaves(direct, reference);
});

test('concurrent direct and replica proposal batches converge on both workers and both main replicas', async () => {
  const bytes = proposalDocument();
  const direct = await openArm(bytes, 6301, true);
  const reference = await openArm(bytes, 6302, false);
  const initialDirect = await direct.client.encodeState();
  const initialReference = await reference.client.encodeState();
  direct.main.applyUpdate(initialReference);
  reference.main.applyUpdate(initialDirect);
  for (const arm of [direct, reference]) {
    await relayout(arm);
    await operate(arm, { kind: 'snapshot' });
  }
  expectCounters(direct, reference, 0);
  await expectEquivalent(direct, reference);
  const x = await propose(direct, [replacement('x', '00000001', 'beta', 'BETA')]);
  const y = await propose(reference, [replacement('y', '00000002', 'epsilon', 'EPSILON')]);
  expectCounters(direct, reference, 1);
  expect(x.changedStories).toEqual(['body']);
  expect(y.changedStories).toEqual(x.changedStories);
  expect(texts(direct.main, 'accepted')).toEqual(['Alpha BETA gamma', 'Delta epsilon', 'User', 'Remote']);
  expect(texts(reference.main, 'accepted')).toEqual(['Alpha beta gamma', 'Delta EPSILON', 'User', 'Remote']);
  for (const update of y.updates) direct.main.applyUpdate(update);
  for (const update of x.updates) reference.main.applyUpdate(update);
  for (const arm of [direct, reference]) await relayout(arm);
  expectCounters(direct, reference, 1);
  const identities = direct.main.paragraphIdentities();
  const segments = direct.main.storySegments('body');
  const revisions = direct.main.listRevisions();
  const vector = direct.main.encodeStateVector();
  for (const arm of [direct, reference]) {
    expect(arm.main.encodeStateVector()).toEqual(vector);
    expect(arm.engine.encodeStateVector()).toEqual(vector);
    expect(arm.main.storySegments('body')).toEqual(segments);
    expect(arm.engine.geometryReader.storySegments('body')).toEqual(segments);
    expect(arm.main.listRevisions()).toEqual(revisions);
    expect(arm.engine.geometryReader.listRevisions()).toEqual(revisions);
    expect(arm.main.paragraphIdentities()).toEqual(identities);
    expect((await arm.client.documentRead({ kind: 'paragraphIdentities' })).value).toEqual(identities);
    for (const view of ['accepted', 'original'] as const) {
      const expected = view === 'accepted'
        ? ['Alpha BETA gamma', 'Delta EPSILON', 'User', 'Remote']
        : TEXTS;
      expect(texts(arm.main, view)).toEqual(expected);
      await expectWorkerParagraphs(arm, view);
    }
  }
  const saved = await expectSaves(direct, reference);
  for (const [index, arm] of [direct, reference].entries()) {
    const workerReplica = await replica(bytes, await arm.client.encodeState(), 6390 + index);
    expect(workerReplica.encodeStateVector()).toEqual(vector);
    expect(workerReplica.storySegments('body')).toEqual(segments);
    expect(workerReplica.paragraphIdentities()).toEqual(identities);
    expect((await saveYrsDocx(workerReplica)).bytes).toEqual(saved);
  }
  expectCounters(direct, reference, 1);
});

test('structural batches fall back and conflicting proposals refuse without advancing the direct counter', async () => {
  const [direct, reference] = await comparison();
  const receipts: DocxEditReceipt[][] = [];
  for (const arm of [direct, reference]) {
    const updates: Uint8Array[] = [];
    const stop = arm.engine.onUpdate((update) => updates.push(update.slice()));
    const steps: DocxEditStep[] = [{
      op: 'insertParagraphs', target: { story: 'body', paraId: '00000003' },
      at: 'end', paragraphs: [{ text: 'Inserted paragraph' }],
    }];
    let result: DocxEditResult;
    try {
      result = arm.engine.proposalEngine.applyEdits({
        expectVersion: arm.engine.proposalEngine.version(), history: 'none', steps,
      });
    } finally {
      stop();
    }
    expect(result).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
    if (!result.ok) throw new Error('expected a structural batch to succeed');
    expect(result.version).toBe(arm.engine.proposalEngine.version());
    receipts.push(result.receipts);
    expect(updates.length).toBeGreaterThan(0);
    arm.adopting = true;
    try {
      for (const update of updates) arm.main.applyHostUpdate(update, ['body']);
    } finally {
      arm.adopting = false;
    }
    await operate(arm, { kind: 'snapshot' });
    expect(texts(arm.main, 'accepted')).toEqual(['Alpha beta gamma', 'Delta epsilon', 'User', 'Inserted paragraph', 'Remote']);
  }
  expect(receipts[0]).toEqual(receipts[1]);
  expectCounters(direct, reference, 0);
  await expectEquivalent(direct, reference);
  await expectSaves(direct, reference);
  await propose(direct, [replacement('owned', '00000001', 'beta', 'BETA')]);
  await propose(reference, [replacement('owned', '00000001', 'beta', 'BETA')]);
  expectCounters(direct, reference, 1);
  const vector = direct.main.encodeStateVector();
  const refused: ResidentProposalReply[] = [];
  for (const arm of [direct, reference]) {
    const reply = await operate(arm, { kind: 'propose', request: {
      expectVersion: arm.engine.proposalEngine.version(),
      proposals: [replacement('conflict', '00000001', 'BETA', 'Other')],
    } });
    expect(reply.result).toMatchObject({ ok: false, failure: { code: 'tracked-revision-conflict' } });
    expect(reply.changedStories).toEqual([]);
    expect(reply.updates).toEqual([]);
    expect(arm.main.encodeStateVector()).toEqual(vector);
    refused.push(reply);
  }
  const [directRefusal, referenceRefusal] = refused;
  if (!directRefusal?.result || directRefusal.result.ok ||
    !referenceRefusal?.result || referenceRefusal.result.ok) {
    throw new Error('expected both conflicting proposals to refuse');
  }
  expect(directRefusal.mirror.proposals).toEqual(referenceRefusal.mirror.proposals);
  expect(directRefusal.result.failure).toEqual(referenceRefusal.result.failure);
  expectCounters(direct, reference, 1);
  await expectEquivalent(direct, reference);
  await expectSaves(direct, reference);
});

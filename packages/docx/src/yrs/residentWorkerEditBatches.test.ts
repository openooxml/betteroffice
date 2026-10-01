import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
  FrameDeltaError,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import { createYrsSession, type DocxEditRequest, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import type { DocxProposalInput, DocxProposalResult, DocxProposalSnapshot } from './proposals';
import { resolveNavigationTarget } from './proposalGeometry';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});

let startWorker: () => InProcessResidentWorker;
const sessions: YrsSession[] = [];
const clients: ResidentEngineWorkerClient[] = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  startWorker = await residentWorkerFactory();
});

afterAll(() => {
  for (const client of clients.splice(0)) client.destroy();
  for (const session of sessions.splice(0)) session.destroy();
});

function accepted(session: YrsSession): string[] {
  const read = session.readParagraphs({ view: 'accepted' });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map((paragraph) => paragraph.text);
}

function frameText(frame: RetainedFrame): string {
  return frame.displayList.pages
    .flatMap((page) => page.primitives)
    .map((primitive) => (primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : ''))
    .join('');
}

test('host batches drain worker input, invalidate the worker once and never adopt stale frames', async () => {
  const main = await createYrsSession({ clientId: 5101 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Seed');
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId, offset: 4 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const forwarded: Uint8Array[] = [];
  let adopting = false;
  main.onUpdate((update) => {
    if (adopting) return;
    forwarded.push(update);
    client.invalidate(update, null);
  });
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(frameText(frame)).toBe('Seed');

  const request = (expectVersion: string): DocxEditRequest => ({
    expectVersion,
    steps: [
      {
        op: 'replaceText',
        target: { kind: 'search', text: 'Seed', within: { kind: 'paragraph', story: 'body', paraId }, view: 'accepted' },
        text: 'Batch',
      },
    ],
  });
  const readBeforeTyping = main.version();
  const typed = await client.applyInput(' typed', main.selection()!, frame.frameEpoch);
  if (!typed.applied) throw new Error('the worker refused resident input');
  adopting = true;
  for (const update of typed.updates) main.applyLocalUpdate(update);
  adopting = false;
  const typedFrame = typed.frame.slice();
  frame = applyFrameDeltaOwned(frame, decodeFrameDelta(typed.frame));
  expect(frameText(frame)).toBe('Seed typed');

  expect(main.applyEdits(request(readBeforeTyping))).toMatchObject({
    ok: false,
    failure: { code: 'stale-version' },
  });
  expect(accepted(main)).toEqual(['Seed typed']);
  expect(forwarded).toHaveLength(0);
  expect(client.isReady()).toBe(true);

  const applied = main.applyEdits(request(main.version()));
  expect(applied).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
  expect(forwarded).toHaveLength(1);
  expect(client.isReady()).toBe(false);
  expect(await client.applyInput('!', main.selection()!, frame.frameEpoch)).toEqual({ applied: false });

  main.layoutDocumentWithRegionsJson(LAYOUT);
  const synced = await client.sync(
    main.residentWorkerSnapshot({
      knownStateVector: client.remoteStateVector(),
      knownFontsRevision: client.syncedFontsRevision(),
    })!,
    '{}',
    frame.frameEpoch
  );
  expect(client.remoteStateVector()).toEqual(main.encodeStateVector());
  frame = applyFrameDeltaOwned(frame, decodeFrameDelta(synced.frame));
  expect(frameText(frame)).toBe('Batch typed');
  let stale: unknown = null;
  try {
    applyFrameDeltaOwned(frame, decodeFrameDelta(typedFrame));
  } catch (error) {
    stale = error;
  }
  expect(stale).toBeInstanceOf(FrameDeltaError);
  expect((stale as FrameDeltaError).code).toBe('stale-frame');
});

test('one resident delete request removes several characters and stops at the story start', async () => {
  const main = await createYrsSession({ clientId: 5102 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Alpha');
  const { secondParaId } = main.splitParagraph({ story: 'body', paraId, offset: 2 });
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId: secondParaId, offset: 1 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(frameText(frame)).toBe('Alpha');

  const remove = async (count: number) => {
    const deleted = await client.applyDelete(
      'backward',
      main.selection()!,
      frame.frameEpoch,
      false,
      false,
      count
    );
    if (!deleted.applied) throw new Error('the worker refused resident deletion');
    for (const update of deleted.updates) main.applyLocalUpdate(update);
    main.setSelection(deleted.selection!.anchor, deleted.selection!.head);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(deleted.frame));
    return deleted.deletedUnits;
  };
  expect(await remove(3)).toBe(3);
  expect(accepted(main)).toEqual(['Aha']);
  expect(frameText(frame)).toBe('Aha');
  expect(await remove(10)).toBe(1);
  expect(accepted(main)).toEqual(['ha']);
  expect(frameText(frame)).toBe('ha');
});

test('one resident delete request merges at most one paragraph', async () => {
  const main = await createYrsSession({ clientId: 5103 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Alpha');
  const { secondParaId } = main.splitParagraph({ story: 'body', paraId, offset: 2 });
  const { secondParaId: lastParaId } = main.splitParagraph({
    story: 'body',
    paraId: secondParaId,
    offset: 0,
  });
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId: lastParaId, offset: 0 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(accepted(main)).toEqual(['Al', '', 'pha']);

  const remove = async (count: number) => {
    const deleted = await client.applyDelete(
      'backward',
      main.selection()!,
      frame.frameEpoch,
      false,
      false,
      count
    );
    if (!deleted.applied) throw new Error('the worker refused resident deletion');
    for (const update of deleted.updates) main.applyLocalUpdate(update);
    main.setSelection(deleted.selection!.anchor, deleted.selection!.head);
    frame = applyFrameDeltaOwned(frame, decodeFrameDelta(deleted.frame));
    return deleted.deletedUnits;
  };
  expect(await remove(3)).toBe(1);
  expect(accepted(main)).toEqual(['Al', 'pha']);
  expect(await remove(3)).toBe(3);
  expect(accepted(main)).toEqual(['pha']);
  expect(frameText(frame)).toBe('pha');
});

test('typing in a table cell goes through the resident worker', async () => {
  const main = await createYrsSession({ clientId: 5104 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Anchor');
  main.insertTable({ story: 'body', paraId, offset: 0 }, 1, 2);
  const cell = main.storyIds().find((story) => story.startsWith('body:'))!;
  const cellParagraph = main.paragraphs(cell)[0]!.paraId;
  main.insertText({ story: cell, paraId: cellParagraph, offset: 0 }, 'Cell');
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: cell, paraId: cellParagraph, offset: 4 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  let frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  expect(frameText(frame)).toContain('Cell');

  const typed = await client.applyInput(' typed', main.selection()!, frame.frameEpoch);
  if (!typed.applied) throw new Error('the worker refused resident input in a cell');
  for (const update of typed.updates) main.applyLocalUpdate(update);
  frame = applyFrameDeltaOwned(frame, decodeFrameDelta(typed.frame));
  expect(frameText(frame)).toContain('Cell typed');
  expect(main.paragraphs(cell)[0]!.text).toBe('Cell typed');
  expect(typed.selection?.head).toMatchObject({ story: cell, offset: 10 });
});

test('a resident delete does not merge a paragraph forward over a table', async () => {
  const main = await createYrsSession({ clientId: 5105 });
  sessions.push(main);
  const { paraId } = main.createStory('body', 'Alpha');
  const { secondParaId } = main.splitParagraph({ story: 'body', paraId, offset: 5 });
  main.insertText({ story: 'body', paraId: secondParaId, offset: 0 }, 'Omega');
  main.insertTable({ story: 'body', paraId: secondParaId, offset: 0 }, 1, 1);
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.layoutDocumentWithRegionsJson(LAYOUT);
  main.setSelection({ story: 'body', paraId, offset: 5 });
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '{}');
  const frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
  const before = accepted(main);

  expect(
    await client.applyDelete('forward', main.selection()!, frame.frameEpoch, false, false)
  ).toEqual({ applied: false });
  expect(accepted(main)).toEqual(before);
  expect(client.isReady()).toBe(true);
});

test('the worker lays a host batch out exactly as the main thread does', async () => {
  const main = await createYrsSession({ clientId: 5110 });
  sessions.push(main);
  const filler = 'lorem ipsum dolor sit amet '.repeat(8);
  const ids = main.loadStories([
    {
      storyId: 'body',
      paragraphs: Array.from({ length: 240 }, (_, index) => ({ text: `${index} ${filler}` })),
    },
  ]).body!;
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.adoptResidentWorkerLayout!(LAYOUT);
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const layoutOptions = () => ({ layoutExtras: '{}', stateVector: main.encodeStateVector() });
  const booted = await client.bootstrap(main.residentWorkerSnapshot()!, '', layoutOptions());

  const sessionId = main.paragraphIdentities().sessionId;
  const proposed = main.proposeChanges({
    expectVersion: main.version(),
    proposals: [0, 90, 200].map((index) => ({
      id: `p${index}`,
      paragraph: { kind: 'session', sessionId, story: 'body', paraId: ids[index]! },
      suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
      op: 'replaceText',
      search: 'dolor',
      replaceWith: 'a replacement long enough to rewrap the paragraph',
      occurrence: 'all',
    })),
  });
  expect(proposed.ok).toBe(true);
  main.adoptResidentWorkerLayout!(LAYOUT);
  const synced = await client.sync(
    main.residentWorkerSnapshot({
      knownStateVector: client.remoteStateVector(),
      knownFontsRevision: client.syncedFontsRevision(),
    })!,
    '',
    booted.caret.frameEpoch,
    false,
    layoutOptions()
  );
  const inWorker = JSON.parse(synced.layoutJson!) as { layout: { pages: unknown[] } };
  expect(inWorker.layout.pages.length).toBeGreaterThan(3);
  expect(inWorker).toEqual(JSON.parse(main.layoutDocumentWithRegionsRetainedJson(LAYOUT)));
});

function proposalDocument(texts = ['Alpha', 'Beta', 'Gamma']): Uint8Array {
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
        `<w:p w14:paraId="0000000${index + 1}"><w:r><w:t>${text}</w:t></w:r></w:p>`
      ).join('') + '<w:sectPr/></w:body></w:document>'
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function proposalResultWithoutVersion(result: DocxProposalResult) {
  if (result.ok) return { ...result, snapshot: { ...result.snapshot, version: '' } };
  return { ...result, version: '' };
}

function paragraphTexts(session: YrsSession, view: 'accepted' | 'original'): string[] {
  const read = session.readParagraphs({ view });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map(({ text }) => text);
}

test('ASCII worker proposals carry verified font requirements and Unicode proposals invalidate them', async () => {
  const bytes = proposalDocument();
  const source = await createYrsSession({ clientId: 5111 });
  sessions.push(source);
  source.openDocx(bytes, true);
  source.registerFont(new Uint8Array(readFileSync(FONT)));
  source.adoptResidentWorkerLayout!(LAYOUT);
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  await client.open(bytes);
  const requirementsJson = await client.fontRequirements(LAYOUT);
  await client.bootstrap({ ...source.residentWorkerSnapshot()!, workerAuthoritative: true }, '', {
    opened: true, layoutExtras: '{}', provisionalPages: 1,
  });
  const initial = await client.proposal({ kind: 'snapshot' });
  const propose = (id: string, paraId: string, search: string, replaceWith: string) =>
    client.proposal({ kind: 'propose', request: { expectVersion: clientVersion, proposals: [{
      id,
      paragraph: { kind: 'persisted', story: { kind: 'body', partUri: '/word/document.xml' }, paraId },
      suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
      op: 'replaceText', search, replaceWith,
    }] } });
  let clientVersion = initial.mirror.version;
  const ascii = await propose('ascii', '00000001', 'Alpha', 'First');
  expect(ascii.result?.ok).toBe(true);
  expect(ascii.fontRequirements).toEqual({ layoutInput: LAYOUT, requirementsJson });
  expect(await client.fontRequirements(LAYOUT)).toBe(requirementsJson);
  clientVersion = ascii.mirror.version;
  const unicode = await propose('unicode', '00000002', 'Beta', '漢字');
  expect(unicode.result?.ok).toBe(true);
  expect(unicode.fontRequirements).toBeUndefined();
  expect(await client.fontRequirements(LAYOUT)).not.toBe(requirementsJson);
  const decided = await client.proposal({ kind: 'setStates', request: {
    expectVersion: unicode.mirror.version,
    expectPreviewVersion: unicode.geometry.previewVersion,
    changes: [{ id: 'unicode', state: 'rejected' }],
  } });
  expect(decided.result?.ok).toBe(true);
  expect(decided.fontRequirements).toBeUndefined();
});

test.each(['', '漢字 Alpha'])('worker font preflight is preserved for an ASCII insertion into %j', async (text) => {
  const bytes = proposalDocument([text, 'Beta', 'Gamma']);
  const source = await createYrsSession({ clientId: 5112 });
  sessions.push(source);
  source.openDocx(bytes, true);
  source.registerFont(new Uint8Array(readFileSync(FONT)));
  source.adoptResidentWorkerLayout!(LAYOUT);
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  await client.open(bytes);
  await client.fontRequirements(LAYOUT);
  await client.bootstrap({ ...source.residentWorkerSnapshot()!, workerAuthoritative: true }, '', {
    opened: true, layoutExtras: '{}', provisionalPages: 1,
  });
  const initial = await client.proposal({ kind: 'snapshot' });
  const inserted = await client.proposal({ kind: 'propose', request: {
    expectVersion: initial.mirror.version,
    proposals: [{
      id: 'insert',
      paragraph: { kind: 'persisted', story: { kind: 'body', partUri: '/word/document.xml' }, paraId: '00000001' },
      suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
      op: 'insertText', at: 'end', text: ' Added',
    }],
  } });
  expect(inserted.result?.ok).toBe(true);
  expect(inserted.fontRequirements).toBeUndefined();
});

test('a worker that has not laid out its document refuses registry operations', async () => {
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  await expect(client.proposal({ kind: 'snapshot' })).rejects.toThrow('has not laid out its document');
  await expect(client.proposal({
    kind: 'withdraw',
    request: { expectVersion: '', ids: [] },
  })).rejects.toThrow('has not laid out its document');
});

test('worker registry operations match a direct registry and hand their records to the main replica', async () => {
  const bytes = proposalDocument();
  const source = await createYrsSession({ clientId: 5106 });
  sessions.push(source);
  source.openDocx(bytes, true);
  source.registerFont(new Uint8Array(readFileSync(FONT)));
  source.adoptResidentWorkerLayout!(LAYOUT);
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  const booted = await client.bootstrap(source.residentWorkerSnapshot()!, '', {
    layoutExtras: '{}',
    provisionalPages: 1,
  });
  const initial = await client.handOver();
  expect(initial.proposals).toEqual({ previewVersion: 0, entries: [] });
  expect(await client.encodeState()).toEqual(initial.state);
  const snapshot = await client.proposal({ kind: 'snapshot' });
  expect(snapshot.result).toBeUndefined();
  expect(snapshot.mirror).toEqual({ version: initial.version, proposals: initial.proposals });
  expect(snapshot.updates).toEqual([]);
  expect(snapshot.changedStories).toEqual([]);

  const proposals: DocxProposalInput[] = source.paragraphIdentities().paragraphs
    .filter((paragraph) => paragraph.session?.story === 'body')
    .map((paragraph, index) => ({
      id: `p${index}`,
      paragraph: paragraph.session!,
      suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
      op: 'replaceText',
      search: ['Alpha', 'Beta', 'Gamma'][index]!,
      replaceWith: ['First', 'Second', 'Third'][index]!,
    }));
  const applied = await client.proposal({
    kind: 'propose',
    request: { expectVersion: snapshot.mirror.version, proposals },
  });
  if (!applied.result?.ok) throw new Error('expected three worker proposals');
  const revisionId = applied.result.snapshot.proposals[0]!.revisionIds[0]!;
  const executionClientId = Number(revisionId.split(':')[0]);
  const main = await createYrsSession({ clientId: executionClientId });
  sessions.push(main);
  main.loadState(initial.state);
  const direct = main.proposeChanges({ expectVersion: main.version(), proposals });
  expect(proposalResultWithoutVersion(applied.result)).toEqual(proposalResultWithoutVersion(direct));
  expect(applied.changedStories).toEqual(['body']);
  expect(applied.updates.length).toBeGreaterThan(0);
  expect(client.remoteStateVector()).toEqual(applied.stateVector);
  expect(client.remoteStateVector()).toEqual(main.encodeStateVector());
  expect(applied.geometry.version).toBe(applied.mirror.version);
  expect(applied.geometry.previewVersion).toBe(0);
  expect(applied.geometry.navigationTargets).toBeUndefined();

  const compareTexts = async () => {
    const fresh = await createYrsSession({ clientId: 5200 + sessions.length });
    const directReplica = await createYrsSession({ clientId: 5300 + sessions.length });
    sessions.push(fresh, directReplica);
    fresh.loadState(await client.encodeState());
    directReplica.loadState(main.encodeState());
    for (const view of ['accepted', 'original'] as const) {
      expect(paragraphTexts(fresh, view)).toEqual(paragraphTexts(directReplica, view));
    }
    return fresh;
  };
  let fresh = await compareTexts();
  expect(paragraphTexts(fresh, 'accepted')).toEqual(['First', 'Second', 'Third']);
  expect(paragraphTexts(fresh, 'original')).toEqual(['Alpha', 'Beta', 'Gamma']);

  const changes = [
    { id: 'p0', state: 'accepted' as const },
    { id: 'p1', state: 'rejected' as const },
  ];
  const decided = await client.proposal({
    kind: 'setStates',
    request: {
      expectVersion: applied.mirror.version,
      expectPreviewVersion: applied.result.snapshot.previewVersion,
      changes,
    },
  });
  const directDecision = main.setProposalStates({
    expectVersion: main.version(),
    expectPreviewVersion: main.getProposals().previewVersion,
    changes,
  });
  expect(proposalResultWithoutVersion(decided.result!)).toEqual(
    proposalResultWithoutVersion(directDecision)
  );
  expect(decided.mirror.version).toBe(applied.mirror.version);
  expect(decided.mirror.proposals.previewVersion).toBe(1);
  expect(decided.geometry.previewVersion).toBe(1);
  expect(decided.updates).toEqual([]);
  expect(decided.changedStories).toEqual([]);
  for (const { id, paragraph } of applied.result.snapshot.proposals) {
    expect(decided.geometry.navigationTargets?.[id]).toEqual(
      resolveNavigationTarget(main, paragraph.story, paragraph.paraId)
    );
    expect(await client.documentRead({
      kind: 'navigationTarget', story: paragraph.story, paraId: paragraph.paraId,
    })).toEqual({ version: decided.mirror.version, value: decided.geometry.navigationTargets![id]! });
  }
  await compareTexts();

  const withdrawn = await client.proposal({
    kind: 'withdraw',
    request: { expectVersion: decided.mirror.version, ids: ['p0'] },
  });
  const directWithdrawal = main.withdrawProposals({ expectVersion: main.version(), ids: ['p0'] });
  expect(proposalResultWithoutVersion(withdrawn.result!)).toEqual(
    proposalResultWithoutVersion(directWithdrawal)
  );
  expect(withdrawn.changedStories).toEqual(['body']);
  expect(withdrawn.geometry.navigationTargets).toBeUndefined();
  expect(withdrawn.updates.length).toBeGreaterThan(0);
  fresh = await compareTexts();
  expect(paragraphTexts(fresh, 'original')).toEqual(['First', 'Beta', 'Gamma']);

  const stale = await client.proposal({
    kind: 'propose',
    request: { expectVersion: initial.version, proposals: [{ ...proposals[2]!, id: 'late' }] },
  });
  const directStale = main.proposeChanges({
    expectVersion: 'stale',
    proposals: [{ ...proposals[2]!, id: 'late' }],
  });
  expect(stale.result).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  expect(proposalResultWithoutVersion(stale.result!)).toEqual(proposalResultWithoutVersion(directStale));
  expect(stale.mirror).toEqual(withdrawn.mirror);
  expect(stale.updates).toEqual([]);
  expect(stale.changedStories).toEqual([]);
  const retried = await client.proposal({
    kind: 'propose',
    request: { expectVersion: initial.version, proposals: proposals.slice(1) },
  });
  const directRetry = main.proposeChanges({ expectVersion: 'stale', proposals: proposals.slice(1) });
  if (!retried.result?.ok) throw new Error('expected an idempotent proposal retry');
  expect(proposalResultWithoutVersion(retried.result)).toEqual(proposalResultWithoutVersion(directRetry));
  expect(retried.mirror).toEqual(withdrawn.mirror);
  expect(retried.updates).toEqual([]);
  expect(retried.changedStories).toEqual([]);
  await compareTexts();

  const identities = await client.documentRead({ kind: 'paragraphIdentities' });
  expect(identities.version).toBe(retried.mirror.version);
  expect(identities.value).toEqual(main.paragraphIdentities());
  expect(identities.value.sessionId).toBe(main.paragraphIdentities().sessionId);
  const anchors = [
    proposals[2]!.paragraph,
    { ...proposals[0]!.paragraph, paraId: 'missing' },
    proposals[1]!.paragraph,
  ];
  const resolved = await client.documentRead({ kind: 'resolveParagraphAnchors', anchors });
  expect(resolved).toEqual({
    version: retried.mirror.version,
    value: { results: anchors.map((anchor) => main.resolveParagraphAnchor(anchor)) },
  });
  for (const view of ['accepted', 'original'] as const) {
    const read = await client.documentRead({ kind: 'readParagraphs', request: { view } });
    expect(read.version).toBe(retried.mirror.version);
    expect({ ...read.value, version: '' }).toEqual({ ...main.readParagraphs({ view }), version: '' });
  }
  const paragraph = proposals[1]!.paragraph;
  if (paragraph.kind !== 'session') throw new Error('expected a session anchor');
  const navigation = await client.documentRead({
    kind: 'navigationTarget',
    story: paragraph.story,
    paraId: paragraph.paraId,
  });
  expect(navigation).toEqual({
    version: retried.mirror.version,
    value: resolveNavigationTarget(main, paragraph.story, paragraph.paraId),
  });
  expect(await client.documentRead({
    kind: 'navigationTarget', story: paragraph.story, paraId: paragraph.paraId,
  })).toEqual(navigation);
  expect(retried.geometry.navigationTargets?.p1).toEqual(navigation.value);

  const handoff = await client.handOver();
  expect(handoff.version).toBe(retried.mirror.version);
  expect(handoff.proposals).toEqual(retried.mirror.proposals);
  main.mirrorWorkerDocument({ version: handoff.version, proposals: handoff.proposals });
  expect(main.getProposals()).toEqual(retried.result.snapshot);
  main.openDocx(bytes, false);
  main.loadState(handoff.state);
  main.mirrorWorkerDocument(null);
  expect(main.getProposals()).toEqual({
    version: main.version(),
    previewVersion: handoff.proposals.previewVersion,
    proposals: handoff.proposals.entries.map(({ record }) => record),
  });
  expect(main.proposeChanges({ expectVersion: 'stale', proposals: proposals.slice(1) }).ok).toBe(true);
  expect(main.setProposalStates({
    expectVersion: main.version(),
    expectPreviewVersion: main.getProposals().previewVersion,
    changes: [{ id: 'p2', state: 'accepted' }],
  }).ok).toBe(true);

  client.rebootstrap();
  await client.bootstrap(source.residentWorkerSnapshot()!, '{}', { frameEpoch: booted.caret.frameEpoch });
  expect((await client.handOver()).proposals).toEqual({ previewVersion: 0, entries: [] });
});

test('worker withdrawal refuses a revision joined by typing with a foreign stamp', async () => {
  const source = await createYrsSession({ clientId: 5109 });
  sessions.push(source);
  source.openDocx(proposalDocument(), true);
  source.registerFont(new Uint8Array(readFileSync(FONT)));
  source.adoptResidentWorkerLayout!(LAYOUT);
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  await client.bootstrap(source.residentWorkerSnapshot()!, '{}');
  const initial = await client.handOver();
  const snapshot = await client.proposal({ kind: 'snapshot' });
  const paragraph = source.paragraphIdentities().paragraphs.find(
    (paragraph) => paragraph.session?.story === 'body'
  )!.session!;
  if (paragraph.kind !== 'session') throw new Error('expected a session anchor');
  const proposal: DocxProposalInput = {
    id: 'shared',
    paragraph,
    suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
    op: 'insertText',
    at: 'end',
    text: ' proposed',
  };
  const applied = await client.proposal({
    kind: 'propose',
    request: { expectVersion: snapshot.mirror.version, proposals: [proposal] },
  });
  if (!applied.result?.ok) throw new Error('expected a worker proposal');
  const record = applied.result.snapshot.proposals[0]!;
  expect(record.revisionIds).toHaveLength(1);
  expect(applied.mirror.proposals.entries.map(({ suggest }) => suggest)).toEqual([proposal.suggest]);
  const main = await createYrsSession({ clientId: Number(record.revisionIds[0]!.split(':')[0]) });
  sessions.push(main);
  main.loadState(initial.state);
  const directApplied = main.proposeChanges({ expectVersion: main.version(), proposals: [proposal] });
  expect(proposalResultWithoutVersion(directApplied)).toEqual(
    proposalResultWithoutVersion(applied.result)
  );
  const beforeTyping = main.encodeStateVector();
  const typed = main.insertText(
    { story: paragraph.story, paraId: paragraph.paraId, offset: 'Alpha proposed'.length },
    ' typed',
    { name: proposal.suggest.author, date: '2026-09-30T00:01:00Z' }
  );
  expect(typed.revisionId).toBe(record.revisionIds[0]);
  client.invalidate(main.encodeStateAsUpdate(beforeTyping), null);
  const joined = await client.proposal({ kind: 'snapshot' });
  expect(joined.mirror.version).not.toBe(applied.mirror.version);
  const beforeWithdrawal = await client.encodeState();
  const refused = await client.proposal({
    kind: 'withdraw',
    request: { expectVersion: joined.mirror.version, ids: ['shared'] },
  });
  const direct = main.withdrawProposals({ expectVersion: main.version(), ids: ['shared'] });
  expect(refused.result).toMatchObject({
    ok: false,
    failure: {
      code: 'tracked-revision-conflict',
      proposalId: 'shared',
      message:
        `proposal shared shares revision ${record.revisionIds[0]}` +
        ' with changes made outside the proposals',
    },
  });
  expect(proposalResultWithoutVersion(refused.result!)).toEqual(
    proposalResultWithoutVersion(direct)
  );
  expect(refused.mirror).toEqual(joined.mirror);
  expect(refused.updates).toEqual([]);
  expect(refused.changedStories).toEqual([]);
  expect(await client.encodeState()).toEqual(beforeWithdrawal);
  const read = await client.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
  expect(read.value).toMatchObject({
    ok: true,
    paragraphs: [{ text: 'Alpha proposed typed' }, { text: 'Beta' }, { text: 'Gamma' }],
  });

  const handoff = await client.handOver();
  expect(handoff.state).toEqual(beforeWithdrawal);
  expect(handoff.proposals).toEqual(joined.mirror.proposals);
  expect(handoff.proposals.entries.map(({ suggest }) => suggest)).toEqual([proposal.suggest]);
  const restored = await createYrsSession({ clientId: 5111 });
  sessions.push(restored);
  restored.loadState(handoff.state);
  restored.mirrorWorkerDocument(structuredClone({
    version: handoff.version,
    proposals: handoff.proposals,
  }));
  restored.mirrorWorkerDocument(null);
  expect(restored.getProposals()).toEqual({
    version: restored.version(),
    previewVersion: handoff.proposals.previewVersion,
    proposals: handoff.proposals.entries.map(({ record }) => record),
  });
  const restoredState = restored.encodeState();
  const restoredProposals = restored.getProposals();
  const events: DocxProposalSnapshot[] = [];
  restored.onProposalChange((snapshot) => events.push(snapshot));
  const restoredRefusal = restored.withdrawProposals({
    expectVersion: restored.version(),
    ids: ['shared'],
  });
  expect(proposalResultWithoutVersion(restoredRefusal)).toEqual(
    proposalResultWithoutVersion(direct)
  );
  expect(restored.encodeState()).toEqual(restoredState);
  expect(restored.getProposals()).toEqual(restoredProposals);
  expect(events).toEqual([]);
  expect(accepted(restored)).toEqual(['Alpha proposed typed', 'Beta', 'Gamma']);
});

test('the main document mirror supplies the worker proposal version precondition', async () => {
  const main = await createYrsSession({ clientId: 5110 });
  sessions.push(main);
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.adoptResidentWorkerLayout!(LAYOUT);
  const localVersion = main.version();
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  await client.open(proposalDocument());
  await client.bootstrap(main.residentWorkerSnapshot()!, '{}', { opened: true });
  const snapshot = await client.proposal({ kind: 'snapshot' });
  main.mirrorWorkerDocument(snapshot.mirror);
  expect(main.workerDocumentMirrored()).toBe(true);
  expect(main.version()).toBe(snapshot.mirror.version);
  expect(main.version()).toBe(snapshot.geometry.version);
  expect(main.version()).not.toBe(localVersion);
  const identities = await client.documentRead({ kind: 'paragraphIdentities' });
  expect(identities.version).toBe(main.version());
  const paragraph = identities.value.paragraphs.find(
    (paragraph) => paragraph.session?.story === 'body'
  )!.session!;
  const proposal: DocxProposalInput = {
    id: 'mirrored',
    paragraph,
    suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
    op: 'insertText',
    at: 'end',
    text: '!',
  };
  const firstVersion = main.version();
  const applied = await client.proposal({
    kind: 'propose',
    request: { expectVersion: firstVersion, proposals: [proposal] },
  });
  expect(applied.result?.ok).toBe(true);
  main.mirrorWorkerDocument(applied.mirror);
  expect(main.version()).toBe(applied.mirror.version);
  expect(main.version()).not.toBe(firstVersion);
  const stale = await client.proposal({
    kind: 'propose',
    request: { expectVersion: firstVersion, proposals: [{ ...proposal, id: 'stale' }] },
  });
  expect(stale.result).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  expect(stale.mirror).toEqual(applied.mirror);
  expect(stale.updates).toEqual([]);
  expect(stale.changedStories).toEqual([]);
  main.mirrorWorkerDocument(stale.mirror);
  const decided = await client.proposal({
    kind: 'setStates',
    request: {
      expectVersion: main.version(),
      expectPreviewVersion: main.getProposals().previewVersion,
      changes: [{ id: 'mirrored', state: 'rejected' }],
    },
  });
  expect(decided.result).toMatchObject({
    ok: true,
    snapshot: { previewVersion: 1, proposals: [{ id: 'mirrored', state: 'rejected' }] },
  });
  main.mirrorWorkerDocument(decided.mirror);
  expect(main.version()).toBe(decided.mirror.version);
  const withdrawn = await client.proposal({
    kind: 'withdraw',
    request: { expectVersion: main.version(), ids: ['mirrored'] },
  });
  expect(withdrawn.result).toMatchObject({ ok: true, snapshot: { proposals: [] } });
  main.mirrorWorkerDocument(withdrawn.mirror);
  expect(main.getProposals().proposals).toEqual([]);
  expect(main.version()).toBe(withdrawn.mirror.version);
  const read = await client.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
  expect(read.value).toMatchObject({
    ok: true,
    paragraphs: [{ text: 'Alpha' }, { text: 'Beta' }, { text: 'Gamma' }],
  });
});

test('worker-authoritative bootstraps and syncs retain the opened document and serve ordered reads', async () => {
  const bytes = proposalDocument();
  const main = await createYrsSession({ clientId: 5108 });
  sessions.push(main);
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.adoptResidentWorkerLayout!(LAYOUT);
  const client = new ResidentEngineWorkerClient(startWorker());
  clients.push(client);
  await client.open(bytes);
  const initial = await client.handOver();
  main.mirrorWorkerDocument({ version: initial.version, proposals: initial.proposals });
  const snapshot = main.residentWorkerSnapshot()!;
  const poisoned = {
    ...snapshot,
    state: new Uint8Array([255]),
    mediaSources: 'invalid JSON',
    selection: {
      anchor: { story: 'missing', paraId: 'missing', offset: 0 },
      head: { story: 'missing', paraId: 'missing', offset: 0 },
    },
  };
  const booted = await client.bootstrap(poisoned, '{}', { opened: true });
  await client.sync(poisoned, '{}', booted.caret.frameEpoch);
  const before = await client.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
  expect(before.value).toMatchObject({
    ok: true,
    paragraphs: [{ text: 'Alpha' }, { text: 'Beta' }, { text: 'Gamma' }],
  });
  const identities = await client.documentRead({ kind: 'paragraphIdentities' });
  const anchor = identities.value.paragraphs.find(
    (paragraph) => paragraph.session?.story === 'body'
  )!.session!;
  const pending = client.proposal({
    kind: 'propose',
    request: {
      expectVersion: before.version,
      proposals: [{
        id: 'ordered',
        paragraph: anchor,
        suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
        op: 'insertText',
        at: 'end',
        text: '!',
      }],
    },
  });
  const read = client.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
  const proposed = await pending;
  expect(proposed.result?.ok).toBe(true);
  expect((await read).value).toMatchObject({
    ok: true,
    paragraphs: [{ text: 'Alpha!' }, { text: 'Beta' }, { text: 'Gamma' }],
  });
  const handoff = await client.handOver();
  main.mirrorWorkerDocument({ version: handoff.version, proposals: handoff.proposals });
  main.openDocx(bytes, false);
  main.loadState(handoff.state);
  main.mirrorWorkerDocument(null);
  const workerIdentities = await client.documentRead({ kind: 'paragraphIdentities' });
  expect(workerIdentities.value).toEqual(main.paragraphIdentities());
  expect(workerIdentities.value.sessionId).toBe(main.paragraphIdentities().sessionId);
  const anchors = [anchor, { ...anchor, paraId: 'missing' }];
  expect((await client.documentRead({ kind: 'resolveParagraphAnchors', anchors })).value).toEqual({
    results: anchors.map((paragraph) => main.resolveParagraphAnchor(paragraph)),
  });
});

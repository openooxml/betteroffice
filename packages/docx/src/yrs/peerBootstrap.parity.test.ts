import { afterAll, afterEach, beforeAll, describe, expect, setSystemTime, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { createHash } from 'node:crypto';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { createEditSession, preloadEditWasm, type EditSession } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { mergeDocxHostMetadata, saveEditorDocument } from './editorSave';
import type { DocxEditStep } from './edits';
import { PeerMetadataError, type YrsDocxHost, type YrsOpeningOptions, type YrsSession } from './index';
import { checkPeerMetadataHeader } from './peerMetadata';
import { saveYrsDocx } from './saveYrsDocx';
import { sessionSourcePackage } from './sessionInternals';
import { wrapSession } from './yrsSessionFacade';
import { yrsToDocument } from './yrsToDocument';

const ROOT = resolve(import.meta.dir, '../../../..');
const FIXTURES = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const owned: YrsSession[] = [];

beforeAll(async () => {
  setSystemTime(new Date('2026-10-02T12:00:00Z'));
  await preloadEditWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
  )));
  await preloadOpcWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')
  )));
});

afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});
afterAll(() => setSystemTime());

function documents(): string[] {
  const roots = process.env.PEER_BOOTSTRAP_DOCS
    ? [resolve(process.env.PEER_BOOTSTRAP_DOCS)]
    : FIXTURES.map((dir) => join(ROOT, dir));
  const found: string[] = [];
  const visit = (path: string) => {
    if (statSync(path).isDirectory()) {
      for (const entry of readdirSync(path).sort()) visit(join(path, entry));
    } else if (path.endsWith('.docx')) {
      found.push(path);
    }
  };
  for (const root of roots) {
    const start = found.length;
    visit(root);
    expect(found.length).toBeGreaterThan(start);
  }
  return found;
}

function session(clientId = 97102): { raw: EditSession; session: YrsSession } {
  const raw = createEditSession(clientId);
  const facade = wrapSession(raw, clientId);
  owned.push(facade);
  return { raw, session: facade };
}

function withoutVersion(reply: unknown): unknown {
  if (typeof reply !== 'object' || reply === null) return reply;
  const { version: _version, baseVersion: _baseVersion, ...content } = reply as Record<string, unknown>;
  return content;
}

function compareReads(a: YrsSession, b: YrsSession): void {
  const before = [a.version(), b.version()];
  expect(a.storyIds()).toEqual(b.storyIds());
  for (const story of a.storyIds()) {
    expect(a.paragraphs(story)).toEqual(b.paragraphs(story));
    expect(a.storyParagraphIds(story)).toEqual(b.storyParagraphIds(story));
    expect(a.headings(story)).toEqual(b.headings(story));
    for (const view of ['accepted', 'original'] as const) {
      expect(withoutVersion(a.readParagraphs({ story, view }))).toEqual(
        withoutVersion(b.readParagraphs({ story, view }))
      );
    }
  }
  const identities = a.paragraphIdentities();
  expect(identities).toEqual(b.paragraphIdentities());
  for (const identity of identities.paragraphs) {
    for (const anchor of [identity.session, identity.source, identity.persisted]) {
      if (anchor) expect(a.resolveParagraphAnchor(anchor)).toEqual(b.resolveParagraphAnchor(anchor));
    }
  }
  expect(a.paragraphSavePlan()).toEqual(b.paragraphSavePlan());
  expect(a.listComments()).toEqual(b.listComments());
  for (const comment of a.listComments()) {
    expect(a.resolveComment(comment.id)).toEqual(b.resolveComment(comment.id));
  }
  for (const revisionView of ['accepted', 'original', 'markup'] as const) {
    expect(withoutVersion(a.exportStructured({ revisionView }))).toEqual(
      withoutVersion(b.exportStructured({ revisionView }))
    );
  }
  const controls = a.listContentControls();
  expect(withoutVersion(controls)).toEqual(withoutVersion(b.listContentControls()));
  if (controls.ok) {
    for (const control of controls.content.controls) {
      const selector = { kind: 'id' as const, controlId: control.controlId };
      expect(withoutVersion(a.findContentControls(selector))).toEqual(
        withoutVersion(b.findContentControls(selector))
      );
      const steps: DocxEditStep[] = [{ op: 'setContentControlText', target: selector, text: 'Parity' }];
      expect(withoutVersion(a.validateEdits({ expectVersion: a.version(), steps }))).toEqual(
        withoutVersion(b.validateEdits({ expectVersion: b.version(), steps }))
      );
    }
  }
  expect([a.version(), b.version()]).toEqual(before);
}

function applyBoth(a: YrsSession, b: YrsSession, steps: DocxEditStep[]): boolean {
  const left = a.validateEdits({ expectVersion: a.version(), steps });
  const right = b.validateEdits({ expectVersion: b.version(), steps });
  expect(withoutVersion(left)).toEqual(withoutVersion(right));
  if (!left.ok) return false;
  const before = [a.version(), b.version()];
  const applied = [a, b].map((peer) => peer.applyEdits({
    expectVersion: peer.version(), history: 'none', steps,
  }));
  expect(withoutVersion(applied[0])).toEqual(withoutVersion(applied[1]));
  for (let at = 0; at < applied.length; at += 1) {
    const result = applied[at]!;
    expect(result.ok).toBe(true);
    if (result.ok && result.applied) expect([a, b][at]!.version()).not.toBe(before[at]);
  }
  return true;
}

function editBoth(a: YrsSession, b: YrsSession): void {
  for (const story of a.storyIds()) {
    const read = a.readParagraphs({ story, view: 'accepted' });
    if (!read.ok) continue;
    for (const paragraph of read.paragraphs) {
      const target = { kind: 'paragraph' as const, story, paraId: paragraph.paraId };
      if (!applyBoth(a, b, [{ op: 'insertText', target, at: 'end', text: ' Peer edit' }])) continue;
      applyBoth(a, b, [{
        op: 'insertParagraphs', target: { story, paraId: paragraph.paraId }, at: 'end',
        paragraphs: [{ text: 'New peer paragraph' }],
      }]);
      applyBoth(a, b, [{ op: 'replaceText', target, text: 'Replaced peer text' }]);
      return;
    }
  }
}

async function compareBytes(
  a: ReturnType<typeof session>, b: ReturnType<typeof session>, host: YrsDocxHost
): Promise<void> {
  const encoder = new TextEncoder();
  expect(encoder.encode(a.raw.materialize_docx()!)).toEqual(encoder.encode(b.raw.materialize_docx()!));
  expect(a.raw.media_sources_json()).toBe(b.raw.media_sources_json());
  expect(a.session.materializeDocx()).toEqual(b.session.materializeDocx());
  expect(new Uint8Array(a.session.materializeDocx()!.originalBuffer!)).toEqual(
    new Uint8Array(b.session.materializeDocx()!.originalBuffer!)
  );
  for (let repeat = 0; repeat < 2; repeat += 1) {
    const savedA = await saveYrsDocx(a.session);
    const savedB = await saveYrsDocx(b.session);
    expect(savedA.bytes).toEqual(savedB.bytes);
    expect(savedA.paragraphs).toEqual(savedB.paragraphs);
    expect(savedA.conflicts).toEqual(savedB.conflicts);
    const editorSave = async (peer: YrsSession): Promise<Uint8Array> => {
      const base = peer.materializeDocx()!;
      const projected = yrsToDocument(peer, mergeDocxHostMetadata(base, host.document));
      return new Uint8Array(await saveEditorDocument(
        peer, projected, projected.package.document.comments ?? []
      ));
    };
    expect(await editorSave(a.session)).toEqual(await editorSave(b.session));
  }
}

function pair(
  bytes: Uint8Array,
  worker: YrsSession,
  host: YrsDocxHost,
  options: YrsOpeningOptions = {}
) {
  const state = worker.encodeState();
  const metadata = worker.encodePeerMetadata();
  const baseline = session();
  const bootstrapped = session();
  baseline.session.setPartialDocument(true);
  bootstrapped.session.setPartialDocument(true);
  const versions = [baseline.session.version(), bootstrapped.session.version()];
  const expectedHost = baseline.session.openDocx(bytes, false, options);
  const expectedEvents: Array<{ origin: string; sourceAvailable: boolean }> = [];
  const stopBaseline = baseline.session.onUpdate((_update, origin) => {
    expectedEvents.push({ origin, sourceAvailable: sessionSourcePackage(baseline.session) !== null });
  });
  baseline.session.loadState(state);
  stopBaseline();
  const scope = bootstrapped.session.mediaScope();
  const observed: Array<{ origin: string; sourceAvailable: boolean }> = [];
  const unsubscribe = bootstrapped.session.onUpdate((_update, origin) => {
    observed.push({ origin, sourceAvailable: sessionSourcePackage(bootstrapped.session) !== null });
  });
  const actualHost = bootstrapped.session.bootstrapPeer(state, metadata, bytes, host, options);
  unsubscribe();
  expect(observed.length).toBeGreaterThan(0);
  expect(observed).toEqual(expectedEvents);
  expect(observed.every((event) => event.sourceAvailable)).toBe(true);
  expect(actualHost.document).toEqual(expectedHost.document);
  expect(actualHost.embeddedFonts).toEqual(expectedHost.embeddedFonts);
  expect(actualHost.fontTableRelationshipsXml).toBe(expectedHost.fontTableRelationshipsXml);
  expect(actualHost.referencedFonts).toEqual(host.referencedFonts);
  expect(actualHost.unusedScriptFonts).toEqual([]);
  expect(actualHost.wholeBody).toBeUndefined();
  expect(bootstrapped.session.mediaScope()).toBe(scope + 1);
  expect(bootstrapped.session.isDisplayOnly()).toBe(false);
  expect(bootstrapped.session.getProposals().proposals).toEqual([]);
  expect(sessionSourcePackage(baseline.session)).toEqual(sessionSourcePackage(bootstrapped.session));
  const length = Number(new DataView(metadata.buffer, metadata.byteOffset, metadata.byteLength).getBigUint64(44, true));
  const wire = JSON.parse(new TextDecoder().decode(metadata.subarray(60, 60 + length))) as {
    media: { parts: unknown[] };
  };
  for (let at = 0; at < wire.media.parts.length; at += 1) {
    const token = `media:${at}`;
    expect(bootstrapped.session.mediaSource(token)).toEqual(baseline.session.mediaSource(token));
    expect(bootstrapped.session.mediaDataUrl(token)).toEqual(baseline.session.mediaDataUrl(token));
  }
  expect(baseline.session.version()).not.toBe(versions[0]);
  expect(bootstrapped.session.version()).not.toBe(versions[1]);
  expect(baseline.session.version()).not.toBe(bootstrapped.session.version());
  for (const [at, peer] of [baseline.session, bootstrapped.session].entries()) {
    const stale = peer.validateEdits({ expectVersion: versions[at]!, steps: [] });
    expect(stale).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
  }
  return { baseline, bootstrapped };
}

describe('peer bootstrap corpus parity', () => {
  for (const mediaTokens of [false, true]) {
    for (const path of documents()) {
      test(`${relative(ROOT, path)} mediaTokens=${mediaTokens}`, async () => {
        const bytes = new Uint8Array(readFileSync(path));
        const worker = session(97101).session;
        const options = { generation: 'peer-bootstrap-parity', mediaTokens };
        const host = worker.openDocx(bytes, true, options);
        const { baseline, bootstrapped } = pair(bytes, worker, host, options);
        compareReads(baseline.session, bootstrapped.session);
        await compareBytes(baseline, bootstrapped, host);
        editBoth(baseline.session, bootstrapped.session);
        compareReads(baseline.session, bootstrapped.session);
        await compareBytes(baseline, bootstrapped, host);
      });
    }
  }
});

function synthetic(): Uint8Array {
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const parts: PartsMap = new Map();
  const set = (name: string, value: string) => parts.set(name, toBytes(value));
  set('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
  set('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rIdDoc" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rIdC" Type="${rel}/comments" Target="comments.xml"/></Relationships>`);
  set('word/document.xml', `<w:document ${ns}><w:body><w:p w14:paraId="0000B001">` +
    '<w:commentRangeStart w:id="1"/><w:r><w:t>Alpha beta</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
    '<w:r><w:commentReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>');
  set('word/comments.xml', `<w:comments ${ns}><w:comment w:id="1" w:author="Source author" w:date="2026-10-01T00:00:00Z">` +
    '<w:p w14:paraId="0000C001"><w:r><w:t>Source comment</w:t></w:r></w:p></w:comment></w:comments>');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function setComment(peer: YrsSession, author: string, body: unknown): void {
  const range = peer.resolveComment('1')[0]!;
  peer.applyRawOps(range.story, [{
    op: 'setComment', id: '1', ranges: [[range.start, range.end]], author,
    date: author ? '2026-10-01T00:00:00Z' : '', body,
  }]);
}

for (const authored of [false, true]) {
  test(`comment baseline excludes worker write history with authored=${authored}`, async () => {
    const bytes = synthetic();
    const worker = session(97101).session;
    const host = worker.openDocx(bytes, true);
    setComment(worker, authored ? 'Worker author' : '', authored ? 'Worker body' : null);
    const { baseline, bootstrapped } = pair(bytes, worker, host);
    compareReads(baseline.session, bootstrapped.session);
    expect(withoutVersion(bootstrapped.session.exportStructured({ revisionView: 'markup', stories: ['comments'] }))).toEqual(
      withoutVersion(baseline.session.exportStructured({ revisionView: 'markup', stories: ['comments'] }))
    );
    await compareBytes(baseline, bootstrapped, host);
    if (!authored) {
      const result = bootstrapped.session.exportStructured({ revisionView: 'markup', stories: ['comments'] });
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.content.stories[0]!.comment?.author).toBe('Source author');
    }
    for (const peer of [baseline.session, bootstrapped.session]) {
      setComment(peer, 'Peer author', 'Peer body');
      peer.loadState(worker.encodeState());
    }
    compareReads(baseline.session, bootstrapped.session);
    expect(withoutVersion(bootstrapped.session.exportStructured({ revisionView: 'markup', stories: ['comments'] }))).toEqual(
      withoutVersion(baseline.session.exportStructured({ revisionView: 'markup', stories: ['comments'] }))
    );
    for (const peer of [baseline.session, bootstrapped.session]) {
      setComment(peer, '', null);
      peer.loadState(worker.encodeState());
    }
    compareReads(baseline.session, bootstrapped.session);
    const baselineResult = baseline.session.exportStructured({ revisionView: 'markup', stories: ['comments'] });
    const result = bootstrapped.session.exportStructured({ revisionView: 'markup', stories: ['comments'] });
    expect(withoutVersion(result)).toEqual(withoutVersion(baselineResult));
    expect(baselineResult.ok).toBe(true);
    if (baselineResult.ok) expect(baselineResult.content.stories[0]!.comment?.author).toBeNull();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.content.stories[0]!.comment?.author).toBeNull();
    await compareBytes(baseline, bootstrapped, host);
  });
}

function snapshot(peer: YrsSession): unknown {
  return {
    state: peer.encodeState(), version: peer.version(), source: sessionSourcePackage(peer),
    materialized: peer.materializeDocx(), stories: peer.storyIds(), comments: peer.listComments(),
    proposals: peer.getProposals(), mediaScope: peer.mediaScope(), mirrored: peer.workerDocumentMirrored(),
    resident: peer.residentWorkerSnapshot(),
  };
}

function rewrite(metadata: Uint8Array, edit: (wire: Record<string, unknown>) => void): Uint8Array {
  const view = new DataView(metadata.buffer, metadata.byteOffset, metadata.byteLength);
  const jsonLength = Number(view.getBigUint64(44, true));
  const wire = JSON.parse(new TextDecoder().decode(metadata.subarray(60, 60 + jsonLength))) as Record<string, unknown>;
  edit(wire);
  const json = new TextEncoder().encode(JSON.stringify(wire));
  const blobs = metadata.subarray(60 + jsonLength);
  const rewritten = new Uint8Array(60 + json.length + blobs.length);
  rewritten.set(metadata.subarray(0, 60));
  new DataView(rewritten.buffer).setBigUint64(44, BigInt(json.length), true);
  rewritten.set(json, 60);
  rewritten.set(blobs, 60 + json.length);
  return rewritten;
}

test('every metadata rejection is atomic and permits a later bootstrap or open', () => {
  const bytes = synthetic();
  const worker = session(97101).session;
  const host = worker.openDocx(bytes, true);
  const state = worker.encodeState();
  const metadata = worker.encodePeerMetadata();
  const cases: Array<[string, Uint8Array, Uint8Array | undefined, Uint8Array]> = [];
  for (const [offset, code] of [[0, 'bad-magic'], [8, 'unsupported-version'], [12, 'shape-mismatch']] as const) {
    const bad = metadata.slice();
    bad[offset] ^= 1;
    cases.push([code, bad, bytes, state]);
  }
  for (const end of [0, 59, 60, metadata.length - 1]) {
    cases.push(['truncated', metadata.slice(0, end), bytes, state]);
  }
  const trailing = new Uint8Array(metadata.length + 1);
  trailing.set(metadata);
  cases.push(['invalid-length', trailing, bytes, state]);
  const overflow = metadata.slice();
  new DataView(overflow.buffer).setBigUint64(44, 0xffffffffffffffffn, true);
  cases.push(['invalid-length', overflow, bytes, state]);
  const json = metadata.slice();
  json[60] = 33;
  cases.push(['invalid-json', json, bytes, state]);
  cases.push(['invalid-json', rewrite(metadata, (wire) => { delete wire.source; }), bytes, state]);
  cases.push(['invalid-metadata', rewrite(metadata, (wire) => {
    const index = wire.index as { parts: Array<{ uri: string }> };
    index.parts[0]!.uri = '';
  }), bytes, state]);
  cases.push(['source-mismatch', rewrite(metadata, (wire) => { wire.source_digest = '0'.repeat(64); }), bytes, state]);
  cases.push(['source-mismatch', rewrite(metadata, (wire) => {
    wire.source_length = bytes.length + 1;
  }), bytes, state]);
  cases.push(['source-mismatch', rewrite(metadata, (wire) => {
    (wire.index as { package_sha256: string }).package_sha256 = '0'.repeat(64);
  }), bytes, state]);
  const wrong = bytes.slice();
  wrong[0] ^= 1;
  cases.push(['source-mismatch', metadata, wrong, state]);
  cases.push(['source-required', metadata, undefined, state]);
  cases.push(['invalid-state', metadata, bytes, new Uint8Array(0)]);
  const encoder = new TextEncoder();
  const invalidParent = new Uint8Array([
    1, 2, 7, 0, 40, 1, 10, ...encoder.encode('other-root'),
    6, ...encoder.encode('parent'), 1, 120,
    40, 0, 7, 0, 5, ...encoder.encode('child'), 1, 120, 0,
  ]);
  cases.push(['invalid-state', metadata, bytes, invalidParent]);
  const invalidPackage = new Uint8Array(bytes.length);
  const invalidDigest = createHash('sha256').update(invalidPackage).digest('hex');
  cases.push(['invalid-package', rewrite(metadata, (wire) => {
    wire.source_digest = invalidDigest;
    (wire.index as { package_sha256: string }).package_sha256 = invalidDigest;
  }), invalidPackage, state]);
  for (const [code, bad, source, update] of cases) {
    for (const recovery of ['bootstrap', 'open'] as const) {
      const peer = session().session;
      peer.setPartialDocument(true);
      const before = snapshot(peer);
      let caught: unknown;
      try { peer.bootstrapPeer(update, bad, source, host); } catch (error) { caught = error; }
      expect(caught).toBeInstanceOf(PeerMetadataError);
      expect(caught).toMatchObject({ code });
      expect(snapshot(peer)).toEqual(before);
      if (recovery === 'bootstrap') peer.bootstrapPeer(state, metadata, bytes, host);
      else { peer.openDocx(bytes, false); peer.loadState(state); }
      expect(peer.paragraphIdentities()).toEqual(worker.paragraphIdentities());
    }
  }
});

test('Rust independently rejects unsupported headers and non-empty documents', () => {
  const bytes = synthetic();
  const worker = session(97101).session;
  const host = worker.openDocx(bytes, true);
  const state = worker.encodeState();
  const metadata = worker.encodePeerMetadata();
  for (const [offset, code] of [[0, 'bad-magic'], [8, 'unsupported-version'], [12, 'shape-mismatch']] as const) {
    const peer = session();
    const bad = metadata.slice();
    bad[offset] ^= 1;
    const before = snapshot(peer.session);
    const raw = peer.raw as EditSession & {
      bootstrap_peer(state: Uint8Array, metadata: Uint8Array, source?: Uint8Array): void;
    };
    let caught: unknown;
    try { raw.bootstrap_peer(state, bad, bytes); } catch (error) { caught = error; }
    expect(caught).toMatchObject({ name: 'PeerMetadataError', code });
    expect(snapshot(peer.session)).toEqual(before);
    peer.session.bootstrapPeer(state, metadata, bytes, host);
  }
  const peer = session().session;
  peer.openDocx(bytes, false);
  const before = snapshot(peer);
  const mismatched = rewrite(metadata, (wire) => { wire.source_digest = '0'.repeat(64); });
  expect(() => peer.bootstrapPeer(state, mismatched, undefined, host)).toThrow(PeerMetadataError);
  expect(snapshot(peer)).toEqual(before);
  peer.bootstrapPeer(state, metadata, undefined, host);
  expect(peer.paragraphIdentities()).toEqual(worker.paragraphIdentities());
  const loaded = snapshot(peer);
  let caught: unknown;
  try { peer.bootstrapPeer(state, metadata, bytes, host); } catch (error) { caught = error; }
  expect(caught).toMatchObject({ name: 'PeerMetadataError', code: 'non-empty-document' });
  expect(snapshot(peer)).toEqual(loaded);
  expect(snapshot(peer)).not.toEqual(before);
  peer.openDocx(bytes, false);
  peer.loadState(state);
  expect(peer.paragraphIdentities()).toEqual(worker.paragraphIdentities());
});

test('metadata export requires an opened source and remains usable after rejection', () => {
  const peer = session().session;
  const before = snapshot(peer);
  let caught: unknown;
  try { peer.encodePeerMetadata(); } catch (error) { caught = error; }
  expect(caught).toMatchObject({ name: 'PeerMetadataError', code: 'missing-source' });
  expect(snapshot(peer)).toEqual(before);
  peer.openDocx(synthetic(), true);
  checkPeerMetadataHeader(peer.encodePeerMetadata());
});

test('header checks handle offset buffers and missing wasm capabilities without mutation', () => {
  const bytes = synthetic();
  const worker = session(97101).session;
  const host = worker.openDocx(bytes, true);
  const state = worker.encodeState();
  const metadata = worker.encodePeerMetadata();
  const offset = new Uint8Array(metadata.length + 9);
  offset.set(metadata, 7);
  checkPeerMetadataHeader(offset.subarray(7, 7 + metadata.length));
  const peer = session();
  const before = snapshot(peer.session);
  Object.defineProperty(peer.raw, 'bootstrap_peer', { value: undefined });
  expect(() => peer.session.bootstrapPeer(state, metadata, bytes, host)).toThrow(PeerMetadataError);
  expect(snapshot(peer.session)).toEqual(before);
  Object.defineProperty(peer.raw, 'encode_peer_metadata', { value: undefined });
  expect(() => peer.session.encodePeerMetadata()).toThrow(PeerMetadataError);
  expect(snapshot(peer.session)).toEqual(before);
});

import { afterAll, afterEach, beforeAll, describe, expect, it, setSystemTime, spyOn } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import {
  DirtyProjectionStories,
  EditorDirtyStories,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  proposalProjectionStories,
  saveEditorDocument,
  serialWorkerSaves,
} from './editorSave';
import type { DocxEditRequest } from './edits';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import type { DocxProposalInput, DocxProposalResult } from './proposals';
import type { ResidentEngineSession } from './residentEngineSession';
import {
  ResidentEngineWorkerClient,
  ResidentWorkerSaveUnavailableError,
} from './residentEngineWorkerClient';
import { yrsToDocument } from './yrsToDocument';

const ROOT = resolve(import.meta.dir, '../../../..');
const FIXTURES = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const TIMEOUT = Number(process.env.WORKER_SAVE_TIMEOUT_MS ?? 120_000);
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SUGGEST = { author: 'Host', date: '2026-09-29T12:00:00Z' };
const FONT = resolve(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const LAYOUT = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
  renderEnv: {},
});

const CLIENT_ID = 123;
let startWorker: (clientId?: number) => InProcessResidentWorker;
beforeAll(async () => {
  setSystemTime(new Date('2026-10-02T12:00:00Z'));
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  );
  await preloadOpcWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')))
  );
  startWorker = await residentWorkerFactory();
});

afterAll(() => {
  setSystemTime();
});

const owned: Array<{ destroy(): void }> = [];
afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

function difference(
  actual: Uint8Array,
  expected: Uint8Array,
  ignored: ReadonlySet<string> = new Set()
): string | null {
  if (actual.length === expected.length && actual.every((byte, at) => byte === expected[at])) return null;
  const a = unzipContainer(actual);
  const b = unzipContainer(expected);
  const names = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort();
  for (const name of names) {
    if (ignored.has(name)) continue;
    const x = a[name];
    const y = b[name];
    if (!x || !y) return `${name}: ${x ? 'only in actual' : 'only in expected'}`;
    if (x.length === y.length && x.every((byte, at) => byte === y[at])) continue;
    let at = 0;
    while (at < x.length && x[at] === y[at]) at += 1;
    if (process.env.WORKER_SAVE_NUMBERS_ONLY === '1') return `part ${names.indexOf(name)} at ${at}`;
    const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes.subarray(Math.max(0, at - 80), at + 80));
    return `${name} at ${at}:\n  actual   ${text(x)}\n  expected ${text(y)}`;
  }
  return null;
}

function documents(): string[] {
  const roots = process.env.WORKER_SAVE_DOCS
    ? [resolve(process.env.WORKER_SAVE_DOCS)]
    : FIXTURES.map((dir) => join(ROOT, dir));
  const found: string[] = [];
  const visit = (path: string) => {
    if (statSync(path).isDirectory()) {
      for (const entry of readdirSync(path).sort()) visit(join(path, entry));
    } else if (path.endsWith('.docx')) {
      found.push(path);
    }
  };
  for (const root of roots) visit(root);
  return found;
}

function paragraph(paraId: string, text: string): string {
  return `<w:p w14:paraId="${paraId}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function synthetic(first = paragraph('0000B001', 'Alpha beta gamma.'), comments?: string): Uint8Array {
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE}.wordprocessingml.header+xml"/>${comments === undefined ? '' : `<Override PartName="/word/comments.xml" ContentType="${OFFICE}.wordprocessingml.comments+xml"/>`}</Types>`
  );
  set(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/>${comments === undefined ? '' : `<Relationship Id="rIdC" Type="${REL}/comments" Target="comments.xml"/>`}</Relationships>`
  );
  if (comments !== undefined) set('word/comments.xml', `<w:comments ${NS}>${comments}</w:comments>`);
  set('word/header1.xml', `<w:hdr ${NS}>${paragraph('0000A001', 'Header text')}</w:hdr>`);
  set(
    'word/document.xml',
    `<w:document ${NS} xmlns:r="${REL}"><w:body>` +
      first +
      paragraph('0000B002', 'Delta epsilon zeta.') +
      paragraph('0000B003', 'Eta theta iota.') +
      `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr></w:body></w:document>`
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function zeroLengthBodyComment(): Uint8Array {
  return synthetic(
    '<w:p w14:paraId="0000B001"><w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/>' +
    '<w:r><w:commentReference w:id="1"/><w:t>Alpha beta gamma.</w:t></w:r></w:p>',
    `<w:comment w:id="1" w:author="Host" w:date="${SUGGEST.date}">` +
    paragraph('0000C001', 'Zero-length comment') + '</w:comment>'
  );
}

async function customXmlNote(): Promise<Uint8Array> {
  const parts: PartsMap = new Map(Object.entries(unzipContainer(synthetic(
    '<w:p w14:paraId="0000B001"><w:r><w:t>Body text</w:t></w:r>' +
    '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
  ))));
  const text = (name: string) => new TextDecoder().decode(parts.get(name)!);
  parts.set('[Content_Types].xml', toBytes(text('[Content_Types].xml').replace(
    '</Types>', `<Override PartName="/word/footnotes.xml" ContentType="${OFFICE}.wordprocessingml.footnotes+xml"/></Types>`
  )));
  parts.set('word/_rels/document.xml.rels', toBytes(text('word/_rels/document.xml.rels').replace(
    '</Relationships>', `<Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`
  )));
  parts.set('word/footnotes.xml', toBytes(
    `<w:footnotes ${NS}><w:footnote w:id="1">` +
    '<w:bookmarkStart w:id="7" w:name="note"/><w:customXml w:element="note" w:uri="urn:example">' +
    paragraph('0000F001', 'Untouched note text') +
    '</w:customXml><w:bookmarkEnd w:id="7"/></w:footnote></w:footnotes>'
  ));
  const source = await repackDocx(await parseDocx(rezipPartsToArrayBuffer(parts), { preloadFonts: false }));
  return new Uint8Array(source);
}

interface Opened {
  bytes: Uint8Array;
  worker: InProcessResidentWorker;
  client: ResidentEngineWorkerClient;
  resident: ResidentEngineSession;
  host: Document;
  replica: Replica;
}

async function open(bytes: Uint8Array): Promise<Opened> {
  const worker = startWorker(CLIENT_ID);
  const client = new ResidentEngineWorkerClient(worker);
  owned.push(client);
  const { hostJson } = await client.open(bytes);
  const session = await createYrsSession({ clientId: CLIENT_ID });
  owned.push(session);
  session.openDocx(bytes.slice(), true);
  const host = decodeDocxHostJson(hostJson, bytes).document;
  return {
    bytes, worker, client, resident: worker.sessions[0]!, host,
    replica: new Replica(session, host),
  };
}

function hostComments(opened: Opened): Comment[] {
  return opened.host.package.document.comments ?? [];
}

async function workerSave(
  opened: Opened,
  comments: Comment[] = hostComments(opened),
  withHost = true,
  withPeer = true
): Promise<Uint8Array<ArrayBuffer>> {
  const { bytes } = await opened.client.save({
    comments,
    ...(withHost ? { host: hostSaveMetadata(opened.host) } : {}),
    ...(withPeer ? { stories: opened.replica.dirtyStories.capture().stories } : {}),
  });
  return new Uint8Array(bytes);
}

const CORE = new Set(['docProps/core.xml']);

async function compareSave(
  opened: Opened,
  comments = hostComments(opened),
  withHost = true,
  withPeer = true
): Promise<Uint8Array<ArrayBuffer>> {
  const saved = await workerSave(opened, comments, withHost, withPeer);
  const expected = await opened.replica.save(comments);
  expect(difference(saved, expected)).toBeNull();
  expect(saved).toEqual(expected);
  return saved;
}

function commentMarkers(bytes: Uint8Array, id: number, part = 'word/document.xml'): string[] {
  const body = new TextDecoder().decode(unzipContainer(bytes)[part]);
  return [...body.matchAll(new RegExp(
    `<w:comment(RangeStart|RangeEnd|Reference)\\b[^>]*\\bw:id="${id}"`, 'g'
  ))].map((match) => match[1]!);
}

async function bootstrap(opened: Opened): Promise<void> {
  const main = opened.replica.session;
  main.registerFont(new Uint8Array(readFileSync(FONT)));
  main.adoptResidentWorkerLayout!(LAYOUT);
  await opened.client.bootstrap(
    { ...main.residentWorkerSnapshot()!, workerAuthoritative: true },
    '{}',
    { opened: true, layoutExtras: '{}' }
  );
}

class Replica {
  readonly dirtyStories = new DirtyProjectionStories();

  constructor(
    readonly session: YrsSession,
    private readonly host: Document,
    private base: Document | null = null
  ) {}

  proposal(call: (session: YrsSession) => DocxProposalResult): DocxProposalResult {
    const known = new Set(this.session.getProposals().proposals.map((proposal) => proposal.id));
    const since = this.session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
    const result = call(this.session);
    if (result.ok) {
      for (const story of proposalProjectionStories(
        known, result, this.session.storiesChangedSince(since).stories
      )) this.dirtyStories.add(story);
    }
    return result;
  }

  async save(comments: Comment[]): Promise<Uint8Array<ArrayBuffer>> {
    const base = this.base ?? this.session.materializeDocx();
    if (!base) throw new Error('the replica has no package');
    const projected = yrsToDocument(
      this.session, mergeDocxHostMetadata(base, this.host),
      this.dirtyStories.projectionOptions()
    );
    this.base = projected;
    const buffer = await saveEditorDocument(this.session, projected, comments);
    projected.originalBuffer = buffer;
    this.dirtyStories.clear();
    return new Uint8Array(buffer);
  }
}

async function peerReplica(opened: Opened): Promise<YrsSession> {
  const peer = await createYrsSession({ clientId: CLIENT_ID + 1 });
  owned.push(peer);
  peer.openDocx(opened.bytes.slice(), false);
  peer.loadState(await opened.client.encodeState());
  opened.replica = new Replica(peer, opened.host);
  return peer;
}

function addComment(opened: Opened, peer: YrsSession, comment: Comment, story = 'body'): void {
  peer.applyRawOps(story, [{
    op: 'setComment', id: String(comment.id), ranges: [[0, 5]],
    author: comment.author, date: comment.date, body: comment.content,
  }]);
  opened.host.package.document.comments = [
    ...hostComments(opened).filter((existing) => existing.id !== comment.id), comment,
  ];
  opened.replica.dirtyStories.add(story);
}

function editHeader(engine: Pick<YrsSession, 'applyEdits' | 'version'>): void {
  expect(engine.applyEdits({
    expectVersion: engine.version(),
    steps: [{
      op: 'insertText', target: { kind: 'paragraph', story: 'hf:rIdH1', paraId: '0000A001' },
      at: 'end', text: ' edited',
    }],
  }).ok).toBe(true);
}

function firstParagraph(
  read: (story: string) => { ok: boolean; paragraphs?: Array<{ paraId: string; text: string }> },
  story: string
): string | null {
  const result = read(story);
  return result.ok ? (result.paragraphs?.find((entry) => entry.text.length > 0)?.paraId ?? null) : null;
}

function insertion(
  read: (story: string) => { ok: boolean; paragraphs?: Array<{ paraId: string; text: string }> },
  host: Document,
  version: string
): DocxEditRequest | null {
  const header = [...(host.package.headers?.keys() ?? [])][0];
  const steps: DocxEditRequest['steps'][number][] = [];
  for (const story of ['body', ...(header === undefined ? [] : [`hf:${header}`])]) {
    const paraId = firstParagraph(read, story);
    if (paraId) {
      steps.push({ op: 'insertText', target: { kind: 'paragraph', story, paraId }, at: 'end', text: ' edited' });
    }
  }
  return steps.length > 0 ? { expectVersion: version, steps } : null;
}

function residentEdit(opened: Opened): boolean {
  const engine = opened.resident.proposalEngine;
  const request = insertion(
    (story) => engine.readParagraphs({ story, view: 'accepted' }),
    opened.host,
    engine.version()
  );
  if (!request) return false;
  const result = engine.applyEdits(request);
  if (!result.ok) throw new Error(`the edit was refused: ${JSON.stringify(result)}`);
  return true;
}

function replicaEdit(replica: Replica, host: Document): void {
  const session = replica.session;
  const request = insertion(
    (story) => session.readParagraphs({ story, view: 'accepted' }),
    host,
    session.version()
  );
  if (!request) return;
  const result = session.applyEdits(request);
  if (!result.ok) throw new Error(`the edit was refused: ${JSON.stringify(result)}`);
  if (result.applied) for (const story of result.changedStories) replica.dirtyStories.add(story);
}

describe('worker save', () => {
  it('integrates saved paragraph ID claims before later worker proposals', async () => {
    const opened = await open(synthetic());
    await bootstrap(opened);
    const peer = await createYrsSession({ clientId: CLIENT_ID + 1 });
    owned.push(peer);
    peer.openDocx(opened.bytes.slice(), false);
    peer.loadState(await opened.client.encodeState());
    peer.splitParagraph({ story: 'body', paraId: '0000B001', offset: 5 });
    opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
    const unsynced = await createYrsSession({ clientId: CLIENT_ID + 2 });
    owned.push(unsynced);
    unsynced.openDocx(opened.bytes.slice(), false);
    unsynced.loadState(peer.encodeState());

    const saved = await opened.client.save({
      comments: [], host: hostSaveMetadata(opened.host), stateVector: peer.encodeStateVector(),
    });
    expect(saved.updates.length).toBeGreaterThan(0);
    for (const update of saved.updates) peer.applyUpdate(update);
    const initial = await opened.client.proposal({ kind: 'snapshot' });
    const proposed = await opened.client.proposal({
      kind: 'propose',
      request: {
        expectVersion: initial.mirror.version,
        proposals: [{
          id: 'after-save',
          paragraph: {
            kind: 'persisted',
            story: { kind: 'body', partUri: '/word/document.xml' },
            paraId: '0000B002',
          },
          suggest: SUGGEST,
          op: 'replaceText',
          search: 'epsilon',
          replaceWith: 'EPSILON',
        }],
      },
    });
    expect(proposed.result?.ok).toBe(true);
    expect(proposed.updates.length).toBeGreaterThan(0);
    for (const update of proposed.updates) {
      peer.applyUpdate(update);
      unsynced.applyUpdate(update);
    }
    expect(peer.encodeStateVector()).toEqual(opened.resident.encodeStateVector());
    expect(peer.readParagraphs({ story: 'body', paraIds: ['0000B002'], view: 'accepted' }))
      .toMatchObject({ ok: true, paragraphs: [{ text: 'Delta EPSILON zeta.' }] });
    expect(unsynced.encodeStateVector()).not.toEqual(opened.resident.encodeStateVector());
  }, TIMEOUT);

  it('returns repairs the editor diff caused in the worker', async () => {
    const opened = await open(synthetic());
    await bootstrap(opened);
    const peer = await createYrsSession({ clientId: CLIENT_ID + 1 });
    owned.push(peer);
    peer.openDocx(opened.bytes.slice(), false);
    peer.loadState(await opened.client.encodeState());
    const at = { story: 'body', paraId: '0000B002', offset: 0 };
    const merged = await opened.client.applyDelete(
      'backward', { anchor: at, head: at }, opened.resident.residentCaretSnapshot().frameEpoch
    );
    expect(merged.applied).toBe(true);
    peer.splitParagraph({ story: 'body', paraId: '0000B001', offset: 5 });
    const repairs: Uint8Array[] = [];
    const unsubscribe = opened.resident.onUpdate((update, origin) => {
      if (origin === 'local') repairs.push(update);
    });
    try {
      opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
      await opened.client.revisionCount();
      expect(repairs.length).toBeGreaterThan(0);
      const identities = opened.resident.paragraphIdentities().paragraphs;
      const keys = identities.flatMap((identity) => identity.session ? [identity.session.paraId] : []);
      expect(new Set(keys).size).toBe(keys.length);
      expect(identities.some((identity) => identity.idOrigin === 'repaired')).toBe(true);
    } finally {
      unsubscribe();
    }
    const saved = await opened.client.save({ comments: [], stateVector: peer.encodeStateVector() });
    expect(saved.updates.length).toBeGreaterThan(0);
    for (const update of saved.updates) peer.applyUpdate(update);
    expect(peer.encodeStateVector()).toEqual(opened.resident.encodeStateVector());
    const initial = await opened.client.proposal({ kind: 'snapshot' });
    const proposed = await opened.client.proposal({
      kind: 'propose',
      request: {
        expectVersion: initial.mirror.version,
        proposals: [{
          id: 'after-repair',
          paragraph: {
            kind: 'persisted', story: { kind: 'body', partUri: '/word/document.xml' }, paraId: '0000B003',
          },
          suggest: SUGGEST, op: 'replaceText', search: 'theta', replaceWith: 'THETA',
        }],
      },
    });
    expect(proposed.result?.ok).toBe(true);
    expect(proposed.updates.length).toBeGreaterThan(0);
    for (const update of proposed.updates) peer.applyUpdate(update);
    expect(peer.encodeStateVector()).toEqual(opened.resident.encodeStateVector());
    expect(peer.readParagraphs({ story: 'body', paraIds: ['0000B003'], view: 'accepted' }))
      .toMatchObject({ ok: true, paragraphs: [{ text: 'Eta THETA iota.' }] });
  }, TIMEOUT);

  it('returns the paragraph IDs a save before bootstrap records', async () => {
    const opened = await open(synthetic());
    const peer = await createYrsSession({ clientId: CLIENT_ID + 1 });
    owned.push(peer);
    peer.openDocx(opened.bytes.slice(), false);
    peer.loadState(await opened.client.encodeState());
    const { secondParaId } = peer.splitParagraph({ story: 'body', paraId: '0000B001', offset: 5 });
    opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
    const beforeSave = peer.encodeStateVector();
    const saved = await opened.client.save({ comments: [], stateVector: beforeSave });
    expect(saved.updates.length).toBeGreaterThan(0);
    for (const update of saved.updates) peer.applyUpdate(update);
    expect(peer.encodeStateVector()).not.toEqual(beforeSave);
    expect(peer.encodeStateVector()).toEqual(opened.resident.encodeStateVector());
    await bootstrap(opened);
    const paragraph = peer.paragraphIdentities().paragraphs
      .find((identity) => identity.session?.paraId === secondParaId)!.persisted!;
    const initial = await opened.client.proposal({ kind: 'snapshot' });
    const proposed = await opened.client.proposal({
      kind: 'propose',
      request: {
        expectVersion: initial.mirror.version,
        proposals: [{
          id: 'after-bootstrap', paragraph,
          suggest: SUGGEST, op: 'replaceText', search: 'beta', replaceWith: 'BETA',
        }],
      },
    });
    expect(proposed.result?.ok).toBe(true);
    expect(proposed.updates.length).toBeGreaterThan(0);
    for (const update of proposed.updates) peer.applyUpdate(update);
    expect(peer.encodeStateVector()).toEqual(opened.resident.encodeStateVector());
    expect(peer.readParagraphs({ story: 'body', paraIds: [secondParaId], view: 'accepted' }))
      .toMatchObject({ ok: true, paragraphs: [{ text: ' BETA gamma.' }] });
  }, TIMEOUT);

  it('matches the editor with no edit, one edit and consecutive saves', async () => {
    const opened = await open(synthetic());
    expect(difference(await compareSave(opened), opened.bytes, CORE)).toBeNull();
    expect(residentEdit(opened)).toBe(true);
    replicaEdit(opened.replica, opened.host);
    await compareSave(opened);
    await compareSave(opened);
    expect(opened.worker.requests.filter((type) => type === 'save')).toHaveLength(3);
  }, TIMEOUT);

  it('uses the opened host when the save omits host metadata', async () => {
    const opened = await open(synthetic());
    await compareSave(opened, hostComments(opened), false);
    expect(residentEdit(opened)).toBe(true);
    replicaEdit(opened.replica, opened.host);
    await compareSave(opened, hostComments(opened), false);
  }, TIMEOUT);

  it('merges the current host metadata across consecutive full saves', async () => {
    const opened = await open(synthetic());
    await compareSave(opened);
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    await compareSave(opened);
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2400 };
    await compareSave(opened);
  }, TIMEOUT);

  it('matches the editor on first and repeated saves of an untouched custom-XML note', async () => {
    const opened = await open(await customXmlNote());
    expect(opened.replica.session.materializeDocx()?.package.footnotes
      ?.find((note) => note.id === 1)?.verbatimXml).toContain('<w:customXml');
    expect(opened.replica.session.storyIds()).toContain('fn:1');
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    await compareSave(opened);
    await compareSave(opened);
  }, TIMEOUT);

  it('keeps untouched note XML across a full save after a body edit', async () => {
    const opened = await open(await customXmlNote());
    expect(opened.replica.session.materializeDocx()?.package.footnotes
      ?.find((note) => note.id === 1)?.verbatimXml).toContain('<w:customXml');
    expect(opened.replica.session.storyIds()).toContain('fn:1');
    for (const engine of [opened.resident.proposalEngine, opened.replica.session]) {
      expect(engine.applyEdits({
        expectVersion: engine.version(),
        steps: [{
          op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: '0000B001' },
          at: 'end', text: ' edited',
        }],
      }).ok).toBe(true);
    }
    opened.replica.dirtyStories.add('body');
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    const save = spyOn(opened.resident, 'save');
    try {
      const saved = await compareSave(opened);
      expect(save).toHaveBeenCalledTimes(1);
      expect(save.mock.calls[0]![4].full).toBe(true);
      expect(unzipContainer(saved)['word/footnotes.xml'])
        .toEqual(unzipContainer(opened.bytes)['word/footnotes.xml']);
    } finally {
      save.mockRestore();
    }
  }, TIMEOUT);

  it('projects the body after a worker comment delete when another story changed', async () => {
    const opened = await open(new Uint8Array(readFileSync(join(
      ROOT, 'packages/docx/src/yrs/__fixtures__/comment-ranges/structure.docx'
    ))));
    const session = opened.replica.session;
    const comments = hostComments(opened);
    const deleted = comments.filter((comment) => comment.parentId === undefined).at(-1);
    if (!deleted) throw new Error('the fixture has no body comment');
    const id = deleted.id;
    const first = await compareSave(opened);
    expect(commentMarkers(first, id)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    const since = opened.resident.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
    const header = session.storyIds().find((story) =>
      story.startsWith('hf:') && opened.host.package.headers?.has(story.slice(3))
    );
    if (!header) throw new Error('the fixture has no header story');
    const paraId = firstParagraph(
      (story) => session.readParagraphs({ story, view: 'accepted' }), header
    );
    if (!paraId) throw new Error('the header has no paragraph');
    const step: DocxEditRequest['steps'][number] = {
      op: 'insertText', target: { kind: 'paragraph', story: header, paraId },
      at: 'end', text: ' edited',
    };
    for (const engine of [opened.resident.proposalEngine, session]) {
      expect(engine.applyEdits({
        expectVersion: engine.version(), steps: [step],
      }).ok).toBe(true);
    }
    opened.replica.dirtyStories.add(header);
    opened.resident.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]);
    expect(opened.resident.storiesChangedSince(since).stories).toEqual(['body', header]);
    session.applyRawOps('body', [{ op: 'removeComment', id: String(id) }]);
    opened.replica.dirtyStories.add('body');
    const remaining = comments.filter((comment) => comment.id !== id && comment.parentId !== id);
    const second = await compareSave(opened, remaining);
    expect(commentMarkers(second, id)).toEqual(['Reference']);
  }, TIMEOUT);

  it('projects only the body on the first save after a peer comment add', async () => {
    const opened = await open(synthetic());
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    const peer = await peerReplica(opened);
    const comment: Comment = {
      id: 1, author: 'Peer', date: SUGGEST.date,
      content: [{
        type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Peer comment' }] }],
      }],
    };
    addComment(opened, peer, comment);
    opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
    const saved = await compareSave(opened);
    expect(commentMarkers(saved, comment.id)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    expect(new TextDecoder().decode(unzipContainer(saved)['word/header1.xml']))
      .toContain('<w:t xml:space="preserve">Header text</w:t>');
  }, TIMEOUT);

  it('preserves unseeded comment markers on the first save after a header and host edit', async () => {
    const opened = await open(zeroLengthBodyComment());
    expect(hostComments(opened).map((comment) => comment.id)).toEqual([1]);
    expect(opened.replica.session.listComments()).toEqual([]);
    for (const engine of [opened.resident.proposalEngine, opened.replica.session]) editHeader(engine);
    opened.replica.dirtyStories.add('hf:rIdH1');
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    const saved = await compareSave(opened);
    expect(commentMarkers(opened.bytes, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    expect(commentMarkers(saved, 1)).toEqual(commentMarkers(opened.bytes, 1));
  }, TIMEOUT);

  it('matches consecutive editor saves for overlapping worker saves after a header and host edit', async () => {
    const opened = await open(zeroLengthBodyComment());
    const peer = await peerReplica(opened);
    expect(hostComments(opened).map((comment) => comment.id)).toEqual([1]);
    expect(peer.listComments()).toEqual([]);
    const editorStories = new EditorDirtyStories();
    editHeader(peer);
    editorStories.add('hf:rIdH1');
    opened.replica.dirtyStories.add('hf:rIdH1');
    opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    const save = serialWorkerSaves(editorStories);
    const run = async (stories: string[]) => {
      const saved = await opened.client.save({
        comments: hostComments(opened), host: hostSaveMetadata(opened.host),
        stories, stateVector: peer.encodeStateVector(),
      });
      editorStories.adoptWorkerSaveUpdates(() => {
        for (const update of saved.updates) peer.applyUpdate(update);
      });
      return new Uint8Array(saved.bytes);
    };
    const [first, second] = await Promise.all([save(run), save(run)]);
    const expectedFirst = await opened.replica.save(hostComments(opened));
    const expectedSecond = await opened.replica.save(hostComments(opened));
    expect(difference(first, expectedFirst)).toBeNull();
    expect(first).toEqual(expectedFirst);
    expect(commentMarkers(first, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    expect(difference(second, expectedSecond)).toBeNull();
    expect(second).toEqual(expectedSecond);
    expect(commentMarkers(expectedSecond, 1)).not.toContain('RangeStart');
    expect(commentMarkers(expectedSecond, 1)).not.toContain('RangeEnd');
    expect(commentMarkers(second, 1)).toEqual(commentMarkers(expectedSecond, 1));
  }, TIMEOUT);

  it('matches the editor bytes for a loaded copy header proposal and host margin change', async () => {
    const opened = await open(zeroLengthBodyComment());
    const peer = await peerReplica(opened);
    expect(hostComments(opened).map((comment) => comment.id)).toEqual([1]);
    expect(peer.listComments()).toEqual([]);
    expect(opened.replica.proposal((session) => session.proposeChanges({
      expectVersion: session.version(),
      proposals: [{
        id: 'header-proposal',
        paragraph: {
          kind: 'persisted', story: { kind: 'header', partUri: '/word/header1.xml' }, paraId: '0000A001',
        },
        suggest: SUGGEST, op: 'replaceText', search: 'Header', replaceWith: 'Changed header',
      }],
    })).ok).toBe(true);
    const stories = opened.replica.dirtyStories.capture().stories;
    expect(stories).toEqual(['hf:rIdH1']);
    opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    const saved = await opened.client.save({
      comments: hostComments(opened), host: hostSaveMetadata(opened.host),
      stories, stateVector: peer.encodeStateVector(),
    });
    for (const update of saved.updates) peer.applyUpdate(update);
    const bytes = new Uint8Array(saved.bytes);
    const expected = await opened.replica.save(hostComments(opened));
    expect(difference(bytes, expected)).toBeNull();
    expect(bytes).toEqual(expected);
    expect(commentMarkers(bytes, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    expect(new TextDecoder().decode(unzipContainer(bytes)['word/header1.xml'])).toContain('Changed header');
    expect(new TextDecoder().decode(unzipContainer(bytes)['word/document.xml'])).toContain('w:top="2000"');
  }, TIMEOUT);

  for (const owner of ['peer', 'worker'] as const) {
    it(`retains zero-length body markers after a ${owner} header comment, header edit and margin change`, async () => {
      const opened = await open(zeroLengthBodyComment());
      const peer = owner === 'peer' ? await peerReplica(opened) : opened.replica.session;
      const comment: Comment = {
        id: 2, author: 'Header', date: SUGGEST.date,
        content: [{
          type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Header comment' }] }],
        }],
      };
      editHeader(peer);
      addComment(opened, peer, comment, 'hf:rIdH1');
      if (owner === 'peer') {
        opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
      } else {
        editHeader(opened.resident.proposalEngine);
        opened.resident.applyRawOps('hf:rIdH1', [{
          op: 'setComment', id: '2', ranges: [[0, 5]],
          author: comment.author, date: comment.date, body: comment.content,
        }]);
      }
      const body = opened.host.package.document;
      body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
      const saved = await compareSave(opened, hostComments(opened), true, owner === 'peer');
      expect(commentMarkers(saved, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      expect(commentMarkers(saved, 2, 'word/header1.xml')).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      expect(new TextDecoder().decode(unzipContainer(saved)['word/document.xml']))
        .toContain('w:top="2000"');
    }, TIMEOUT);

    it(`matches the editor for a ${owner} header comment add and delete with no other edit`, async () => {
      const opened = await open(synthetic());
      const peer = owner === 'peer' ? await peerReplica(opened) : opened.replica.session;
      const comment: Comment = {
        id: 1, author: 'Header', date: SUGGEST.date,
        content: [{
          type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Header comment' }] }],
        }],
      };
      addComment(opened, peer, comment, 'hf:rIdH1');
      if (owner === 'peer') {
        opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
      } else {
        opened.resident.applyRawOps('hf:rIdH1', [{
          op: 'setComment', id: '1', ranges: [[0, 5]],
          author: comment.author, date: comment.date, body: comment.content,
        }]);
      }
      const first = await compareSave(opened, [comment], true, owner === 'peer');
      expect(commentMarkers(first, 1, 'word/header1.xml')).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      peer.applyRawOps('hf:rIdH1', [{ op: 'removeComment', id: '1' }]);
      opened.replica.dirtyStories.add('hf:rIdH1');
      opened.host.package.document.comments = [];
      if (owner === 'peer') {
        opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
      } else {
        opened.resident.applyRawOps('hf:rIdH1', [{ op: 'removeComment', id: '1' }]);
      }
      const second = await compareSave(opened, [], true, owner === 'peer');
      expect(commentMarkers(second, 1, 'word/header1.xml')).toEqual([]);
      await compareSave(opened, [], true, owner === 'peer');
    }, TIMEOUT);
  }

  it('matches the editor for a body comment delete with no other edit', async () => {
    const opened = await open(synthetic());
    const peer = await peerReplica(opened);
    const comment: Comment = {
      id: 1, author: 'Body', date: SUGGEST.date,
      content: [{
        type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Body comment' }] }],
      }],
    };
    addComment(opened, peer, comment);
    opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
    expect(commentMarkers(await compareSave(opened), 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    peer.applyRawOps('body', [{ op: 'removeComment', id: '1' }]);
    opened.replica.dirtyStories.add('body');
    opened.host.package.document.comments = [];
    opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
    expect(commentMarkers(await compareSave(opened), 1)).toEqual([]);
    await compareSave(opened);
  }, TIMEOUT);

  it('matches the editor for a viewer worker body comment delete with no other edit', async () => {
    const opened = await open(synthetic(
      '<w:p w14:paraId="0000B001"><w:commentRangeStart w:id="1"/>' +
      '<w:r><w:t>Alpha</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
      '<w:r><w:commentReference w:id="1"/></w:r></w:p>',
      `<w:comment w:id="1" w:author="Host" w:date="${SUGGEST.date}">` +
      paragraph('0000C001', 'Body comment') + '</w:comment>'
    ));
    await bootstrap(opened);
    const removed = await opened.client.proposal({ kind: 'removeComment', id: '1' });
    expect(removed.projectionStories).toEqual(['body']);
    opened.replica.session.applyRawOps('body', [{ op: 'removeComment', id: '1' }]);
    opened.replica.dirtyStories.add('body');
    opened.host.package.document.comments = [];
    const saved = await compareSave(opened, [], true, false);
    expect(commentMarkers(saved, 1)).not.toContain('RangeStart');
    expect(commentMarkers(saved, 1)).not.toContain('RangeEnd');
    await compareSave(opened, [], true, false);
  }, TIMEOUT);

  it('saves a viewer worker proposal with the same marked stories as the editor', async () => {
    const opened = await open(zeroLengthBodyComment());
    await bootstrap(opened);
    const proposals: DocxProposalInput[] = [{
      id: 'header-proposal',
      paragraph: {
        kind: 'persisted', story: { kind: 'header', partUri: '/word/header1.xml' }, paraId: '0000A001',
      },
      suggest: SUGGEST, op: 'replaceText', search: 'Header', replaceWith: 'Changed header',
    }];
    const initial = await opened.client.proposal({ kind: 'snapshot' });
    const proposed = await opened.client.proposal({
      kind: 'propose', request: { expectVersion: initial.mirror.version, proposals },
    });
    expect(proposed.result?.ok).toBe(true);
    expect(proposed.projectionStories).toEqual(['hf:rIdH1']);
    expect(opened.replica.proposal((session) => session.proposeChanges({
      expectVersion: session.version(), proposals,
    })).ok).toBe(true);
    const body = opened.host.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
    const first = await compareSave(opened, hostComments(opened), true, false);
    expect(commentMarkers(first, 1)).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
    expect(new TextDecoder().decode(unzipContainer(first)['word/header1.xml'])).toContain('Changed header');
    await compareSave(opened, hostComments(opened), true, false);
  }, TIMEOUT);

  for (const operation of ['add', 'delete', 'delete with unchanged host comments'] as const) {
    it(`projects the body after a peer comment ${operation} when another story changed`, async () => {
      const opened = await open(synthetic());
      const body = opened.host.package.document;
      body.finalSectionProperties = { ...body.finalSectionProperties, marginTop: 2000 };
      const peer = await peerReplica(opened);
      const comment: Comment = {
        id: 1,
        author: 'Peer',
        date: SUGGEST.date,
        content: [{
          type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Peer comment' }] }],
        }],
      };
      if (operation !== 'add') {
        addComment(opened, peer, comment);
        opened.client.invalidate(peer.encodeStateAsUpdate(opened.client.remoteStateVector()!), null);
      }
      const first = await compareSave(opened);
      expect(commentMarkers(first, comment.id)).toEqual(
        operation === 'add' ? [] : ['RangeStart', 'RangeEnd', 'Reference']
      );
      editHeader(peer);
      opened.replica.dirtyStories.add('hf:rIdH1');
      if (operation === 'add') {
        addComment(opened, peer, comment);
      } else {
        peer.applyRawOps('body', [{ op: 'removeComment', id: String(comment.id) }]);
        if (operation === 'delete') opened.host.package.document.comments = [];
        opened.replica.dirtyStories.add('body');
      }
      const update = peer.encodeStateAsUpdate(opened.client.remoteStateVector()!);
      opened.client.invalidate(update, null);
      await opened.client.encodeState();
      const second = await compareSave(opened);
      expect(commentMarkers(second, comment.id)).toEqual(
        operation === 'add' ? ['RangeStart', 'RangeEnd', 'Reference'] : []
      );
      if (operation === 'delete with unchanged host comments') {
        expect(hostComments(opened)).toEqual([comment]);
        const savedComments = unzipContainer(second)['word/comments.xml'];
        expect(new TextDecoder().decode(savedComments)).toMatch(/<w:comment\b[^>]*\bw:id="1"/);
      }
      opened.client.invalidate(update, null);
      await opened.client.encodeState();
      const third = await compareSave(opened);
      expect(commentMarkers(third, comment.id)).toEqual(commentMarkers(second, comment.id));
    }, TIMEOUT);
  }

  it('writes a host comment reply range once across two saves', async () => {
    const opened = await open(new Uint8Array(readFileSync(join(
      ROOT, 'packages/docx/src/yrs/__fixtures__/comment-ranges/structure.docx'
    ))));
    const comments = hostComments(opened);
    const parent = comments.find((comment) => comment.parentId === undefined);
    if (!parent) throw new Error('the fixture has no comment');
    const reply: Comment = {
      id: Math.max(...comments.map((comment) => comment.id)) + 1,
      author: 'Host',
      date: SUGGEST.date,
      content: parent.content,
      parentId: parent.id,
    };
    const first = await compareSave(opened, [...comments, reply]);
    const second = await compareSave(opened, [...comments, reply]);
    const body = (bytes: Uint8Array) =>
      new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
    expect(body(second)).toBe(body(first));
    expect(body(second).split(`<w:commentRangeStart w:id="${reply.id}"/>`)).toHaveLength(2);
  }, TIMEOUT);

  for (const state of ['accepted', 'rejected'] as const) {
    it(`matches the editor after host proposals are all ${state}`, async () => {
      const opened = await open(synthetic());
      await bootstrap(opened);
      const inputs: DocxProposalInput[] = [
        ['a', '0000B001', 'beta', 'BETA'],
        ['b', '0000B002', 'epsilon', 'EPSILON'],
        ['c', '0000B003', 'theta', 'THETA'],
      ].map(([id, paraId, search, replaceWith]) => ({
        id: id!,
        paragraph: {
          kind: 'persisted',
          story: { kind: 'body', partUri: '/word/document.xml' },
          paraId: paraId!,
        },
        suggest: SUGGEST,
        op: 'replaceText',
        search: search!,
        replaceWith: replaceWith!,
      }));
      const main = opened.replica.session;
      const initial = await opened.client.proposal({ kind: 'snapshot' });
      const proposed = await opened.client.proposal({
        kind: 'propose',
        request: { expectVersion: initial.mirror.version, proposals: inputs },
      });
      expect(proposed.result?.ok).toBe(true);
      expect(opened.replica.proposal((session) => session.proposeChanges({
        expectVersion: session.version(), proposals: inputs,
      })).ok).toBe(true);
      const changes = inputs.map(({ id }) => ({ id, state }));
      const decided = await opened.client.proposal({
        kind: 'setStates',
        request: {
          expectVersion: proposed.mirror.version,
          expectPreviewVersion: proposed.mirror.proposals.previewVersion,
          changes,
        },
      });
      expect(decided.result?.ok).toBe(true);
      expect(opened.replica.proposal((session) => session.setProposalStates({
        expectVersion: main.version(),
        expectPreviewVersion: main.getProposals().previewVersion,
        changes,
      })).ok).toBe(true);
      const ids = inputs.map(({ id }) => id);
      const withdrawn = await opened.client.proposal({
        kind: 'withdraw',
        request: { expectVersion: decided.mirror.version, ids },
      });
      expect(withdrawn.result?.ok).toBe(true);
      expect(opened.replica.proposal((session) => session.withdrawProposals({
        expectVersion: session.version(), ids,
      })).ok).toBe(true);
      const first = await compareSave(opened);
      const second = await compareSave(opened);
      if (state === 'rejected') {
        expect(difference(first, opened.bytes, CORE)).toBeNull();
        expect(difference(second, opened.bytes, CORE)).toBeNull();
      } else {
        const body = new TextDecoder().decode(unzipContainer(first)['word/document.xml']);
        for (const text of ['BETA', 'EPSILON', 'THETA']) expect(body).toContain(text);
      }
    }, TIMEOUT);
  }

  it('saves worker input like the editor', async () => {
    const opened = await open(synthetic());
    await bootstrap(opened);
    const caret = { story: 'body', paraId: '0000B003', offset: 3 };
    const selection = { anchor: caret, head: caret };
    const typed = await opened.client.applyInput(' typed', selection, opened.client.answeredFrame());
    expect(typed.applied).toBe(true);
    opened.replica.session.insertText(caret, ' typed');
    opened.replica.dirtyStories.add(caret.story);
    await compareSave(opened);
    await compareSave(opened);
  }, TIMEOUT);

  it('keeps a raw inline offset across edits and consecutive saves', async () => {
    const opened = await open(synthetic(
      '<w:p w14:paraId="0000B001"><w:r><w:t>abcdefghij</w:t></w:r>' +
      '<x:mark xmlns:x="urn:example"/><w:r><w:t>klm</w:t></w:r></w:p>'
    ));
    const edit = (step: DocxEditRequest['steps'][number]) => {
      for (const engine of [opened.replica.session, opened.resident.proposalEngine]) {
        expect(engine.applyEdits({
          expectVersion: engine.version(), history: 'none', steps: [step],
        }).ok).toBe(true);
      }
      opened.replica.dirtyStories.add('body');
    };
    const rawOffset = (bytes: Uint8Array) => {
      const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
      const at = xml.indexOf('<x:mark');
      expect(at).toBeGreaterThanOrEqual(0);
      return [...xml.slice(0, at).matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)]
        .reduce((length, match) => length + match[1]!.length, 0);
    };
    expect(rawOffset(opened.bytes)).toBe(10);
    edit({
      op: 'deleteText',
      target: {
        kind: 'range', story: 'body',
        start: { paraId: '0000B001', offset: 3 },
        end: { paraId: '0000B001', offset: 13 },
        view: 'accepted',
      },
    });
    expect(rawOffset(await compareSave(opened))).toBe(3);
    edit({
      op: 'insertText',
      target: { kind: 'paragraph', story: 'body', paraId: '0000B001' },
      at: 'end', text: 'ABCDEFGHIJ',
    });
    expect(rawOffset(await compareSave(opened))).toBe(3);
  }, TIMEOUT);

  for (const path of documents()) {
    it(`matches the editor before and after an edit: ${relative(ROOT, path)}`, async () => {
      const opened = await open(new Uint8Array(readFileSync(path)));
      expect(difference(await compareSave(opened), opened.bytes, CORE)).toBeNull();
      if (!residentEdit(opened)) return;
      replicaEdit(opened.replica, opened.host);
      await compareSave(opened);
      await compareSave(opened);
    }, TIMEOUT);
  }
});

describe('worker save availability', () => {
  it('rejects a save before opening a document', async () => {
    const client = new ResidentEngineWorkerClient(startWorker(CLIENT_ID));
    owned.push(client);
    await expect(client.save({ comments: [] })).rejects.toBeInstanceOf(ResidentWorkerSaveUnavailableError);
    expect(client.hasFailed()).toBe(false);
  });

  it('rejects a save after a peer-only bootstrap replaces an opened document', async () => {
    const opened = await open(synthetic());
    await bootstrap(opened);
    await compareSave(opened);
    opened.client.rebootstrap();
    await opened.client.bootstrap(opened.replica.session.residentWorkerSnapshot()!, '{}', {
      layoutExtras: '{}',
    });
    await expect(opened.client.save({ comments: [] }))
      .rejects.toBeInstanceOf(ResidentWorkerSaveUnavailableError);
    expect(opened.client.hasFailed()).toBe(false);
  }, TIMEOUT);

  it('rejects a save while previewing, then saves the whole opened document', async () => {
    const bytes = synthetic();
    const worker = startWorker(CLIENT_ID);
    const client = new ResidentEngineWorkerClient(worker);
    owned.push(client);
    expect(await client.openPreview(bytes, 1)).not.toBeNull();
    const unavailable = client.save({ comments: [] });
    await expect(unavailable).rejects.toBeInstanceOf(ResidentWorkerSaveUnavailableError);
    await expect(unavailable).rejects.toThrow('still opening');
    client.rebootstrap();
    await client.open(bytes);
    const main = await createYrsSession({ clientId: CLIENT_ID });
    owned.push(main);
    const host = main.openDocx(bytes.slice(), true).document;
    const replica = new Replica(main, host);
    const saved = new Uint8Array((await client.save({ comments: [] })).bytes);
    expect(difference(saved, await replica.save([]))).toBeNull();
    expect(difference(saved, bytes, CORE)).toBeNull();
  }, TIMEOUT);

  it('resets save history when destroy is followed by open', async () => {
    const opened = await open(synthetic());
    expect(residentEdit(opened)).toBe(true);
    await workerSave(opened);
    opened.worker.postMessage({ id: -1, type: 'destroy' });
    const client = new ResidentEngineWorkerClient(opened.worker);
    owned.push(client);
    const bytes = synthetic(paragraph('0000B001', 'A new document.'));
    const { hostJson } = await client.open(bytes);
    const main = await createYrsSession({ clientId: CLIENT_ID });
    owned.push(main);
    main.openDocx(bytes.slice(), true);
    const replica = new Replica(main, decodeDocxHostJson(hostJson, bytes).document);
    const saved = new Uint8Array((await client.save({ comments: [] })).bytes);
    expect(difference(saved, await replica.save([]))).toBeNull();
    expect(difference(saved, bytes, CORE)).toBeNull();
  }, TIMEOUT);
});

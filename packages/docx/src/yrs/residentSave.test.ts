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
  dirtyProjectionStory,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  saveEditorDocument,
} from './editorSave';
import type { DocxEditRequest } from './edits';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import type { DocxProposalInput } from './proposals';
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

function synthetic(first = paragraph('0000B001', 'Alpha beta gamma.')): Uint8Array {
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE}.wordprocessingml.header+xml"/></Types>`
  );
  set(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="${REL}/header" Target="header1.xml"/></Relationships>`
  );
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
  withHost = true
): Promise<Uint8Array<ArrayBuffer>> {
  const { bytes } = await opened.client.save({
    comments,
    ...(withHost ? { host: hostSaveMetadata(opened.host) } : {}),
  });
  return new Uint8Array(bytes);
}

const CORE = new Set(['docProps/core.xml']);

async function compareSave(
  opened: Opened,
  comments = hostComments(opened),
  withHost = true
): Promise<Uint8Array<ArrayBuffer>> {
  const saved = await workerSave(opened, comments, withHost);
  expect(difference(saved, await opened.replica.save(comments))).toBeNull();
  return saved;
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
  private revision: number;

  constructor(
    readonly session: YrsSession,
    private readonly host: Document,
    private base: Document | null = null
  ) {
    this.revision = session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
  }

  async save(comments: Comment[]): Promise<Uint8Array<ArrayBuffer>> {
    const base = this.base ?? this.session.materializeDocx();
    if (!base) throw new Error('the replica has no package');
    const storyIds = new Set(this.session.storiesChangedSince(this.revision).stories.map(dirtyProjectionStory));
    const projected = yrsToDocument(
      this.session, mergeDocxHostMetadata(base, this.host),
      storyIds.size > 0 ? { storyIds } : undefined
    );
    this.base = projected;
    const buffer = await saveEditorDocument(this.session, projected, comments);
    projected.originalBuffer = buffer;
    this.revision = this.session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
    return new Uint8Array(buffer);
  }
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

  it('keeps untouched note XML across a full save after a body edit', async () => {
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
    const opened = await open(new Uint8Array(source));
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
      expect(main.proposeChanges({ expectVersion: main.version(), proposals: inputs }).ok).toBe(true);
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
      expect(main.setProposalStates({
        expectVersion: main.version(),
        expectPreviewVersion: main.getProposals().previewVersion,
        changes,
      }).ok).toBe(true);
      const ids = inputs.map(({ id }) => id);
      const withdrawn = await opened.client.proposal({
        kind: 'withdraw',
        request: { expectVersion: decided.mirror.version, ids },
      });
      expect(withdrawn.result?.ok).toBe(true);
      expect(main.withdrawProposals({ expectVersion: main.version(), ids }).ok).toBe(true);
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
    opened.client.destroy();
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

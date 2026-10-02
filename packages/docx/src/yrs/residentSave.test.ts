import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../layout/render/frameDelta';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import {
  adoptEditorSave,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  saveEditorDocument,
} from './editorSave';
import type { DocxEditRequest } from './edits';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { createProposalRegistry, type DocxProposalInput } from './proposals';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import type { ResidentSaveRecord } from './residentSave';
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

let startWorker: () => InProcessResidentWorker;
beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  );
  startWorker = await residentWorkerFactory();
});

const owned: Array<{ destroy(): void }> = [];
afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

/** Where two packages first differ; null when they are byte-identical. */
function difference(actual: Uint8Array, expected: Uint8Array): string | null {
  const length = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < length && actual[at] === expected[at]) at += 1;
  return at === length && actual.length === expected.length
    ? null
    : `byte ${at} of ${actual.length} vs ${expected.length}`;
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

function synthetic(): Uint8Array {
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
      paragraph('0000B001', 'Alpha beta gamma.') +
      paragraph('0000B002', 'Delta epsilon zeta.') +
      paragraph('0000B003', 'Eta theta iota.') +
      `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr></w:body></w:document>`
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

interface Opened {
  bytes: Uint8Array;
  resident: ResidentEngineSession;
  hostJson: string;
  host: Document;
  record: ResidentSaveRecord;
}

async function open(bytes: Uint8Array): Promise<Opened> {
  const resident = await createResidentEngineSession();
  owned.push(resident);
  const hostJson = resident.openDocx(bytes.slice());
  return { bytes, resident, hostJson, host: decodeDocxHostJson(hostJson, bytes).document, record: { full: false } };
}

/** The worker's save, its inputs crossing as they do in a message. */
async function workerSave(opened: Opened, comments: Comment[] = hostComments(opened)): Promise<Uint8Array<ArrayBuffer>> {
  return new Uint8Array(
    await opened.resident.save(
      opened.bytes.slice(),
      opened.hostJson,
      structuredClone(hostSaveMetadata(opened.host)),
      structuredClone(comments),
      opened.record
    )
  );
}

/** The main-thread replica the editor hydrates from the worker and saves. */
class Replica {
  private base: Document | null = null;

  constructor(
    readonly session: YrsSession,
    private readonly host: Document
  ) {}

  async save(comments: Comment[]): Promise<Uint8Array<ArrayBuffer>> {
    const base = this.base ?? this.session.materializeDocx();
    if (!base) throw new Error('the replica has no package');
    const projected = yrsToDocument(this.session, mergeDocxHostMetadata(base, this.host));
    this.base = projected;
    const buffer = await saveEditorDocument(this.session, projected, comments);
    projected.originalBuffer = buffer;
    return new Uint8Array(buffer);
  }
}

async function hydrate(opened: Opened): Promise<Replica> {
  const session = await createYrsSession();
  owned.push(session);
  session.openDocx(opened.bytes.slice(), false);
  session.loadState(opened.resident.encodeState());
  return new Replica(session, opened.host);
}

function hostComments(opened: Opened): Comment[] {
  return opened.host.package.document.comments ?? [];
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
  it('equals the replica save on the synthetic fixture: no edit, edits, two saves', async () => {
    const opened = await open(synthetic());
    const replica = await hydrate(opened);
    expect(difference(await workerSave(opened), await replica.save(hostComments(opened)))).toBeNull();
    expect(residentEdit(opened)).toBe(true);
    replicaEdit(replica, opened.host);
    expect(difference(await workerSave(opened), await replica.save(hostComments(opened)))).toBeNull();
    expect(difference(await workerSave(opened), await replica.save(hostComments(opened)))).toBeNull();
  });

  it('equals the replica save after host proposals are accepted and rejected', async () => {
    const opened = await open(synthetic());
    const registry = createProposalRegistry(opened.resident.proposalEngine);
    const replace = (id: string, paraId: string, search: string, replaceWith: string): DocxProposalInput => ({
      id,
      paragraph: { kind: 'persisted', story: { partUri: '/word/document.xml', kind: 'body' }, paraId },
      suggest: SUGGEST,
      op: 'replaceText',
      search,
      replaceWith,
    });
    const proposed = registry.propose({
      expectVersion: opened.resident.proposalEngine.version(),
      proposals: [
        replace('a', '0000B001', 'beta', 'BETA'),
        replace('b', '0000B002', 'epsilon', 'EPSILON'),
        replace('c', '0000B003', 'theta', 'THETA'),
      ],
    });
    expect(proposed.ok).toBe(true);
    const decided = registry.setStates({
      expectVersion: opened.resident.proposalEngine.version(),
      expectPreviewVersion: registry.snapshot().previewVersion,
      changes: [
        { id: 'a', state: 'accepted' },
        { id: 'b', state: 'rejected' },
      ],
    });
    expect(decided.ok).toBe(true);
    const replica = await hydrate(opened);
    expect(difference(await workerSave(opened), await replica.save(hostComments(opened)))).toBeNull();
  });

  it('equals the replica save with the host comments and a reply', async () => {
    const bytes = new Uint8Array(
      readFileSync(join(ROOT, 'packages/docx/src/yrs/__fixtures__/comment-ranges/structure.docx'))
    );
    const opened = await open(bytes);
    const comments = hostComments(opened);
    const parent = comments.find((comment) => comment.parentId === undefined);
    if (!parent) throw new Error('the fixture has no comment');
    const reply: Comment = {
      id: Math.max(...comments.map((comment) => comment.id)) + 1,
      author: 'Host',
      date: '2026-09-29T12:00:00Z',
      content: parent.content,
      parentId: parent.id,
    };
    const replica = await hydrate(opened);
    expect(
      difference(await workerSave(opened, [...comments, reply]), await replica.save([...comments, reply]))
    ).toBeNull();
    expect(
      difference(await workerSave(opened, [...comments, reply]), await replica.save([...comments, reply]))
    ).toBeNull();
  });

  it('equals a replica that hydrates after the worker saved', async () => {
    const reference = await open(synthetic());
    const replica = await hydrate(reference);
    replicaEdit(replica, reference.host);
    const first = await replica.save(hostComments(reference));
    replicaEdit(replica, reference.host);
    const second = await replica.save(hostComments(reference));

    const opened = await open(synthetic());
    expect(residentEdit(opened)).toBe(true);
    expect(difference(await workerSave(opened), first)).toBeNull();
    const late = await hydrate(opened);
    const record = opened.record;
    adoptEditorSave(late.session, record.full ? { full: true } : { full: false, saved: record.original });
    const materialized = late.session.materializeDocx();
    if (!materialized) throw new Error('the replica has no package');
    replicaEdit(late, opened.host);
    const projected = yrsToDocument(
      late.session,
      mergeDocxHostMetadata({ ...materialized, originalBuffer: record.original }, opened.host)
    );
    const lateSave = new Uint8Array(await saveEditorDocument(late.session, projected, hostComments(opened)));
    expect(difference(lateSave, second)).toBeNull();
  });

  for (const path of documents()) {
    const name = relative(ROOT, path);
    it(
      `equals the replica save, before and after an edit: ${name}`,
      async () => {
        const opened = await open(new Uint8Array(readFileSync(path)));
        const replica = await hydrate(opened);
        expect(difference(await workerSave(opened), await replica.save(hostComments(opened)))).toBeNull();
        if (!residentEdit(opened)) return;
        replicaEdit(replica, opened.host);
        expect(difference(await workerSave(opened), await replica.save(hostComments(opened)))).toBeNull();
      },
      TIMEOUT
    );
  }
});

describe('worker save request', () => {
  it('saves typed input and decided proposals as the replica saves them', async () => {
    const bytes = synthetic();
    const main = await createYrsSession();
    owned.push(main);
    main.openDocx(bytes, true);
    main.registerFont(new Uint8Array(readFileSync(FONT)));
    main.adoptResidentWorkerLayout!(LAYOUT);
    const client = new ResidentEngineWorkerClient(startWorker());
    owned.push(client);
    const { hostJson } = await client.open(bytes);
    const host = decodeDocxHostJson(hostJson, bytes).document;
    const booted = await client.bootstrap(
      { ...main.residentWorkerSnapshot()!, workerAuthoritative: true },
      '{}',
      { opened: true, layoutExtras: '{}' }
    );
    const frame = applyFrameDeltaOwned(null, decodeFrameDelta(booted.frame));
    const caret = { story: 'body', paraId: '0000B003', offset: 3 };
    const typed = await client.applyInput(' typed', { anchor: caret, head: caret }, frame.frameEpoch);
    expect(typed.applied).toBe(true);
    const initial = await client.proposal({ kind: 'snapshot' });
    const proposed = await client.proposal({
      kind: 'propose',
      request: {
        expectVersion: initial.mirror.version,
        proposals: [
          {
            id: 'a',
            paragraph: { kind: 'persisted', story: { kind: 'body', partUri: '/word/document.xml' }, paraId: '0000B001' },
            suggest: SUGGEST,
            op: 'replaceText',
            search: 'beta',
            replaceWith: 'BETA',
          },
        ],
      },
    });
    expect(proposed.result?.ok).toBe(true);
    const accepted = await client.proposal({
      kind: 'setStates',
      request: {
        expectVersion: proposed.mirror.version,
        expectPreviewVersion: proposed.mirror.proposals.previewVersion,
        changes: [{ id: 'a', state: 'accepted' }],
      },
    });
    expect(accepted.result?.ok).toBe(true);

    const replicaSession = await createYrsSession();
    owned.push(replicaSession);
    replicaSession.openDocx(bytes, false);
    replicaSession.loadState(await client.encodeState());
    const replica = new Replica(replicaSession, host);
    const request = { source: bytes, hostJson, host: hostSaveMetadata(host), comments: [] };
    const first = await client.save(request);
    expect(difference(new Uint8Array(first.bytes), await replica.save([]))).toBeNull();
    const second = await client.save(request);
    expect(difference(new Uint8Array(second.bytes), await replica.save([]))).toBeNull();
    expect(second.full).toBe(first.full);
  });

  it('refuses a save without an opened document', async () => {
    const client = new ResidentEngineWorkerClient(startWorker());
    owned.push(client);
    await expect(
      client.save({ source: synthetic(), hostJson: '{}', host: hostSaveMetadata({ package: { document: { content: [] } } }), comments: [] })
    ).rejects.toThrow('has no opened document');
  });
});

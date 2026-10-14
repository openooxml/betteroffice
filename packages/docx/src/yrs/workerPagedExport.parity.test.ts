import { beforeAll, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { basename, join, relative, resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import type { DocxLayoutMap, DocxPageExportOptions, DocxPagedStructuredContent } from './pagedExport';
import type { DocxParagraphIdentitySnapshot } from './paragraphIdentity';
import {
  createProposalRegistry,
  type DocxProposalRequest,
  type DocxProposalResult,
  type DocxProposalSession,
} from './proposals';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import type { DocxStorySelection } from './readTypes';
import type { DocxExportResult } from './structuredExport';

const ROOT = resolve(import.meta.dir, '../../../..');
const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = new Uint8Array(
  readFileSync(join(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))
);
const FIXTURES = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const ALL: DocxStorySelection[] = ['body', 'headers', 'footers', 'footnotes', 'endnotes', 'comments'];
const OPTIONS: DocxPageExportOptions[] = [
  { revisionView: 'markup', stories: ALL },
  { revisionView: 'markup', stories: ALL, includeGeometry: true },
  { revisionView: 'accepted', stories: ALL },
  { revisionView: 'original', stories: ALL },
];
const SUGGEST = { author: 'Reviewer', date: '2026-01-01T00:00:00Z' };
const TIMEOUT = Number(process.env.PAGED_EXPORT_PARITY_TIMEOUT_MS ?? 60_000);
type PagedExport = DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>;

let startWorker: (clientId?: number) => InProcessResidentWorker;
beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  startWorker = await residentWorkerFactory();
});

function documents(): string[] {
  const roots = process.env.PAGED_EXPORT_PARITY_DOCS
    ? [resolve(process.env.PAGED_EXPORT_PARITY_DOCS)]
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
  return found.sort();
}

function normalize(reply: PagedExport): PagedExport {
  reply.version = '<version>';
  if (reply.ok) {
    reply.content.layout.documentVersion = '<version>';
    reply.content.layout.layoutVersion = '<layoutVersion>';
    reply.content.layout.provenance.layoutEpoch = '<layoutEpoch>';
  }
  return reply;
}

function difference(worker: unknown, replica: unknown, path = '$'): string | undefined {
  if (JSON.stringify(worker) === JSON.stringify(replica)) return undefined;
  if (worker && replica && typeof worker === 'object' && typeof replica === 'object') {
    const left = worker as Record<string, unknown>;
    const right = replica as Record<string, unknown>;
    const keys = Object.keys(left);
    if (JSON.stringify(keys) !== JSON.stringify(Object.keys(right))) {
      return `${path} keys: worker=${JSON.stringify(keys)}, replica=${JSON.stringify(Object.keys(right))}`;
    }
    for (const key of keys) {
      const found = difference(left[key], right[key], `${path}.${key}`);
      if (found) return found;
    }
  }
  return `${path}: worker=${JSON.stringify(worker)}, replica=${JSON.stringify(replica)}`;
}

function workerExport(
  worker: ResidentEngineSession,
  options: DocxPageExportOptions,
  request: string
): PagedExport {
  const reply = JSON.parse(worker.exportStructuredWithPagesJson(options, request)) as PagedExport;
  expect(reply.version).toBe(worker.proposalEngine.version());
  if (reply.ok) {
    const layout = reply.content.layout;
    expect(layout.documentVersion).toBe(reply.version);
    expect(layout.layoutVersion.startsWith(`${layout.documentVersion}:`)).toBe(true);
  }
  return reply;
}

function leaves(
  worker: unknown,
  main: unknown,
  path = '$',
  found: Array<[string, unknown, unknown]> = []
): Array<[string, unknown, unknown]> {
  if (JSON.stringify(worker) === JSON.stringify(main)) return found;
  if (worker && main && typeof worker === 'object' && typeof main === 'object') {
    const left = worker as Record<string, unknown>;
    const right = main as Record<string, unknown>;
    for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
      leaves(left[key], right[key], `${path}.${key}`, found);
    }
  } else {
    found.push([path, worker, main]);
  }
  return found;
}

const PLACEMENT = new Set(['id', 'nodeId', 'blockId', 'anchor']);
const SOURCE_PLACED = new Set(['break', 'unsupported']);

/**
 * `value` without what only a seeded session can place from the package: source breaks and
 * omitted inline source content, the ids and anchors their placement shifts, and the split of
 * the text runs around them.
 */
function withoutSourcePlacement(value: unknown): unknown {
  if (Array.isArray(value)) {
    const kept = value.filter((item) =>
      !(item && typeof item === 'object' && SOURCE_PLACED.has(String((item as { kind?: unknown }).kind))));
    const merged: unknown[] = [];
    for (const item of kept.map(withoutSourcePlacement)) {
      const previous = merged.at(-1) as Record<string, unknown> | undefined;
      const current = item as Record<string, unknown> | null;
      if (previous?.kind === 'text' && current?.kind === 'text' &&
        JSON.stringify({ ...previous, text: '' }) === JSON.stringify({ ...current, text: '' })) {
        merged[merged.length - 1] = { ...previous, text: `${previous.text}${current.text}` };
      } else {
        merged.push(item);
      }
    }
    return merged;
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([key]) => !PLACEMENT.has(key))
        .map(([key, item]) => [key, withoutSourcePlacement(item)])
    );
  }
  return value;
}

function sourceProvenanceGap(worker: PagedExport, main: PagedExport): boolean {
  if (!worker.ok || !main.ok) return false;
  const reported = (reply: typeof main) => reply.content.structured.diagnostics.some(
    ({ code }) => code === 'provenance-unavailable'
  );
  if (reported(worker)) return false;
  if (reported(main)) {
    const comparable = ({ content }: typeof main) => JSON.stringify({
      structured: withoutSourcePlacement({ ...content.structured, diagnostics: [] }),
      pages: content.layout.pages,
      occurrences: withoutSourcePlacement(content.layout.occurrences),
    });
    return comparable(worker) === comparable(main);
  }
  return leaves(worker, main).every(([path]) => path === '$.content.layout.exportFingerprint');
}

/**
 * Exports `worker` and `main` with every option set; equal after the session-local tokens.
 * With `tolerateProvenanceGap`, an export may differ by {@link sourceProvenanceGap} only; such
 * exports are counted.
 */
function compare(
  worker: ResidentEngineSession,
  main: YrsSession,
  request: string,
  tolerateProvenanceGap = false
): number {
  const failures: string[] = [];
  let gaps = 0;
  for (const options of OPTIONS) {
    const workerReply = normalize(workerExport(worker, options, request));
    const mainReply = normalize(
      JSON.parse(JSON.stringify(main.exportStructuredWithPagesFor(options, request))) as PagedExport
    );
    const workerJson = JSON.stringify(workerReply);
    const mainJson = JSON.stringify(mainReply);
    if (workerJson === mainJson) continue;
    if (tolerateProvenanceGap && sourceProvenanceGap(workerReply, mainReply)) {
      gaps += 1;
      continue;
    }
    failures.push(`${JSON.stringify(options)} ${difference(JSON.parse(workerJson), JSON.parse(mainJson))}`);
  }
  expect(failures.join('\n')).toBe('');
  return gaps;
}

let nextClientId = 97300;

function laidOutRequest(session: Pick<YrsSession, 'registerFont' | 'layoutFontRequirementsJson'>, document: Parameters<typeof buildResidentRegionLayoutRequest>[0]): string {
  const font = session.registerFont(FONT);
  const layout = buildResidentRegionLayoutRequest(document, 24, {});
  const requirements = JSON.parse(
    session.layoutFontRequirementsJson(JSON.stringify(layout))
  ) as Array<{ key: string }>;
  layout.measurement = {
    fontChains: Object.fromEntries(requirements.map(({ key }) => [key, [font]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  return JSON.stringify(layout);
}

/** The main-thread session today's viewer exports from: opened unseeded and loaded from the worker. */
async function replicaOf(bytes: Uint8Array, worker: ResidentEngineSession, request: string): Promise<YrsSession> {
  const replica = await createYrsSession({ clientId: nextClientId++ });
  replica.openDocx(bytes, false);
  replica.loadState(worker.encodeState());
  replica.registerFont(FONT);
  replica.layoutDocumentWithRegionsRetainedJson(request);
  return replica;
}

/** A main-thread session seeded from the same bytes, as a session opened on the main thread is. */
async function seededOf(bytes: Uint8Array, clientId: number, request: string): Promise<YrsSession> {
  const seeded = await createYrsSession({ clientId });
  seeded.openDocx(bytes, true);
  seeded.registerFont(FONT);
  seeded.layoutDocumentWithRegionsRetainedJson(request);
  return seeded;
}

function propose(
  session: Pick<DocxProposalSession, 'readParagraphs'>,
  identities: DocxParagraphIdentitySnapshot,
  expectVersion: string,
  submit: (request: DocxProposalRequest) => DocxProposalResult
): void {
  const paragraphs = identities.paragraphs
    .flatMap(({ session: anchor }) => anchor?.story === 'body' ? [anchor] : []);
  const read = session.readParagraphs({
    story: 'body', paraIds: paragraphs.map(({ paraId }) => paraId), view: 'accepted',
  });
  if (!read.ok) throw new Error(read.failure.message);
  const first = read.paragraphs.find(({ text }) => text.length > 1);
  if (!first) throw new Error('expected a paragraph to replace');
  const replacement = paragraphs.find(({ paraId }) => paraId === first.paraId)!;
  const insertion = paragraphs.find(({ paraId }) => paraId !== first.paraId);
  if (!insertion) throw new Error('expected a second paragraph to insert into');
  const result = submit({
    expectVersion,
    proposals: [
      {
        id: 'replacement', paragraph: replacement, suggest: SUGGEST,
        op: 'replaceText', search: first.text, replaceWith: 'Replacement text',
      },
      {
        id: 'insertion', paragraph: insertion, suggest: SUGGEST,
        op: 'insertText', at: 'end', text: ' Added text',
      },
  ],
  });
  if (!result.ok) throw new Error(result.failure.message);
  expect(result.snapshot.proposals).toHaveLength(2);
  expect(result.snapshot.proposals.every(
    ({ changed, revisionIds }) => changed && revisionIds.length > 0
  )).toBe(true);
}

/** Compares the worker's export with a seeded main-thread session (strict) and with today's replica. */
async function parity(bytes: Uint8Array, afterProposal = false): Promise<number> {
  const clientId = nextClientId++;
  const worker = await createResidentEngineSession(undefined, clientId);
  const sessions: YrsSession[] = [];
  try {
    const host = decodeDocxHostJson(worker.openDocx(bytes), bytes);
    const request = laidOutRequest(worker, host.document);
    worker.layoutDocumentWithRegionsRetainedJson(request);
    const seeded = await seededOf(bytes, clientId, request);
    sessions.push(seeded);
    if (afterProposal) {
      const exported = workerExport(worker, OPTIONS[0]!, request);
      if (!exported.ok) throw new Error(exported.failure.message);
      const registry = createProposalRegistry(worker.proposalEngine);
      try {
        propose(worker.proposalEngine, worker.paragraphIdentities(), exported.version, (request) => registry.propose(request));
      } finally {
        registry.destroy();
      }
      propose(seeded, seeded.paragraphIdentities(), seeded.version(), (request) => seeded.proposeChanges(request));
      worker.layoutDocumentWithRegionsRetainedJson(request);
      seeded.layoutDocumentWithRegionsRetainedJson(request);
      const relaid = workerExport(worker, OPTIONS[0]!, request);
      if (!relaid.ok) throw new Error(relaid.failure.message);
      const layoutVersion = relaid.content.layout.layoutVersion;
      expect(workerExport(worker, { ...OPTIONS[0]!, expectLayoutVersion: layoutVersion }, request).ok).toBe(true);
      expect(workerExport(worker, {
        ...OPTIONS[0]!, expectLayoutVersion: `${layoutVersion}:different`,
      }, request)).toMatchObject({ ok: false, failure: { code: 'stale-layout' } });
    }
    compare(worker, seeded, request);
    const replica = await replicaOf(bytes, worker, request);
    sessions.push(replica);
    return compare(worker, replica, request, true);
  } finally {
    for (const session of sessions) session.destroy();
    worker.destroy();
  }
}

function synthetic(sourceBreaks = false): Uint8Array {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const parts = new Map<string, Uint8Array>();
  const xml = (name: string, content: string) => parts.set(name, toBytes(content));
  const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const section = (index: number) => `<w:headerReference w:type="default" r:id="header${index}"/><w:footerReference w:type="default" r:id="footer${index}"/><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>`;
  const stories: Array<[string, string]> = [
    ['comments', 'comments'],
    ['header1', 'header'], ['footer1', 'footer'],
    ['header2', 'header'], ['footer2', 'footer'],
  ];
  xml('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${stories.map(([name, kind]) => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`).join('')}</Types>`);
  xml('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  xml('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${stories.map(([name, kind]) => `<Relationship Id="${name}" Type="${R}/${kind}" Target="${name}.xml"/>`).join('')}<Relationship Id="image" Type="${R}/image" Target="media/pixel.png"/></Relationships>`);
  for (const [name, kind] of stories) {
    if (kind !== 'comments') {
      const tag = kind === 'header' ? 'hdr' : 'ftr';
      xml(`word/${name}.xml`, `<w:${tag} xmlns:w="${W}">${paragraph(name)}</w:${tag}>`);
    }
  }
  xml('word/comments.xml', `<w:comments xmlns:w="${W}">${[0, 1].map((id) => `<w:comment w:id="${id}" w:author="Reviewer" w:initials="R" w:date="2026-01-01T00:00:00Z">${paragraph(`Comment ${id + 1}`)}</w:comment>`).join('')}</w:comments>`);
  const comments = [0, 1].map((id) => `<w:commentRangeStart w:id="${id}"/><w:r><w:t>Marked ${id + 1} </w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`).join('');
  const revisions = `<w:ins w:id="10" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Inserted </w:t></w:r></w:ins><w:del w:id="11" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Deleted </w:delText></w:r></w:del>`;
  const image = '<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Pixel"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Pixel"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${paragraph('Merged cell')}</w:tc></w:tr><w:tr><w:tc>${paragraph('Left cell')}</w:tc><w:tc>${paragraph('Right cell')}</w:tc></w:tr></w:tbl>`;
  const breaks = sourceBreaks ? `<w:p><w:r><w:t>Before source break</w:t><w:br w:type="page"/><w:t>After source break</w:t></w:r></w:p>${paragraph('Closing paragraph')}` : '';
  xml('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${paragraph('First paragraph')}<w:p>${comments}${revisions}${image}</w:p>${table}<w:p><w:pPr><w:sectPr>${section(1)}</w:sectPr></w:pPr><w:r><w:t>Section boundary</w:t></w:r></w:p>${paragraph('Second section')}${breaks}<w:sectPr>${section(2)}</w:sectPr></w:body></w:document>`);
  parts.set('word/media/pixel.png', new Uint8Array(Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aD1sAAAAASUVORK5CYII=',
    'base64'
  )));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

test('editor-style worker export after hydration and peer edits strictly preserves source page breaks', async () => {
  const bytes = synthetic(true);
  const clientId = nextClientId++;
  const port = startWorker(clientId);
  const client = new ResidentEngineWorkerClient(port);
  const sessions: YrsSession[] = [];
  try {
    const opened = await client.open(bytes);
    const host = decodeDocxHostJson(opened.hostJson, bytes);
    const resident = port.sessions[0]!;
    const request = laidOutRequest(resident, host.document);
    resident.layoutDocumentWithRegionsRetainedJson(request);
    const seeded = await seededOf(bytes, clientId, request);
    const peer = await replicaOf(bytes, resident, request);
    sessions.push(seeded, peer);
    expect(compare(resident, peer, request, true)).toBeGreaterThan(0);
    const first = peer.paragraphs('body')[0]!;
    expect(first.text).toBe('First paragraph');
    const edited = peer.applyEdits({ expectVersion: peer.version(), steps: [{
      op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: first.paraId }, at: 'end', text: ' Peer edit',
    }] });
    if (!edited.ok) throw new Error(edited.failure.message);
    expect(peer.paragraphs('body')[0]!.text).toBe('First paragraph Peer edit');
    propose(peer, peer.paragraphIdentities(), peer.version(), (request) => peer.proposeChanges(request));
    const update = peer.encodeStateAsUpdate(client.remoteStateVector()!);
    const capturedVector = peer.encodeStateVector();
    const acknowledged = await client.syncUpdate(update, capturedVector);
    if (acknowledged.repair) peer.applyLocalUpdate(acknowledged.repair);
    seeded.applyUpdate(update);
    if (acknowledged.repair) seeded.applyUpdate(acknowledged.repair);
    resident.layoutDocumentWithRegionsRetainedJson(request);
    seeded.layoutDocumentWithRegionsRetainedJson(request);
    for (const options of OPTIONS) {
      const read = await client.documentReadAt({ kind: 'exportStructuredWithPages', options, currentRequest: request }, acknowledged.version);
      if (read.status !== 'ok') throw new Error('The acknowledged export was superseded');
      const exported = JSON.parse(read.value) as PagedExport;
      const source = seeded.exportStructuredWithPagesFor(options, request);
      expect(normalize(exported)).toEqual(normalize(source));
      expect(exported.ok).toBe(options.revisionView === 'markup');
      if (!exported.ok) {
        expect(exported.failure.code).toBe('unsupported-revision-layout');
        continue;
      }
      expect(JSON.stringify(exported.content.structured)).toContain('"breakType":"page"');
      expect(exported.content.structured.diagnostics.some(({ code }) => code === 'provenance-unavailable')).toBe(false);
      const comments = exported.content.structured.stories.filter(({ kind }) => kind === 'comment');
      expect(comments).toHaveLength(2);
      expect(comments.every(({ comment }) => comment?.author === 'Reviewer' && comment.date === '2026-01-01T00:00:00Z')).toBe(true);
    }
    expect(port.requests).toContain('syncUpdate');
  } finally {
    for (const session of sessions) session.destroy();
    client.destroy();
  }
}, TIMEOUT);

test('a deletion-only peer diff reaches the worker despite unchanged state vectors', async () => {
  const bytes = synthetic(true);
  const port = startWorker(nextClientId++);
  const client = new ResidentEngineWorkerClient(port);
  let peer: YrsSession | null = null;
  try {
    const opened = await client.open(bytes);
    const resident = port.sessions[0]!;
    const host = decodeDocxHostJson(opened.hostJson, bytes);
    const request = laidOutRequest(resident, host.document);
    peer = await replicaOf(bytes, resident, request);
    const before = peer.encodeStateVector();
    peer.applyRawOps('body', [{ op: 'delete', index: 0, len: 1 }]);
    expect(peer.encodeStateVector()).toEqual(before);
    const acknowledged = await client.syncUpdate(peer.encodeStateAsUpdate(client.remoteStateVector()!), before);
    if (acknowledged.repair) peer.applyLocalUpdate(acknowledged.repair);
    expect(resident.proposalEngine.readParagraphs({ story: 'body', paraIds: [peer.paragraphs('body')[0]!.paraId], view: 'accepted' })).toMatchObject({
      ok: true, paragraphs: [{ text: 'irst paragraph' }],
    });
    resident.layoutDocumentWithRegionsRetainedJson(request);
    const read = await client.documentReadAt({ kind: 'exportStructuredWithPages', options: OPTIONS[0]!, currentRequest: request }, acknowledged.version);
    expect(read.status).toBe('ok');
    if (read.status !== 'ok') throw new Error('The deletion was superseded');
    expect(JSON.parse(read.value)).toMatchObject({ ok: true });
    expect(read.value).toContain('"breakType":"page"');
  } finally {
    peer?.destroy();
    client.destroy();
  }
}, TIMEOUT);

describe("the worker's paged export equals a main-thread session's", () => {
  for (const file of documents()) {
    const name = relative(process.env.PAGED_EXPORT_PARITY_DOCS ?? ROOT, file);
    test(name, async () => { await parity(new Uint8Array(readFileSync(file))); }, TIMEOUT);
    if (basename(file) === 'pages.docx') {
      test(`${name} after proposals`, async () => {
        expect(await parity(new Uint8Array(readFileSync(file)), true)).toBe(0);
      }, TIMEOUT);
    }
  }
  test('synthetic stories, revisions, comments, merged cell and image', async () => {
    expect(await parity(synthetic())).toBe(0);
  }, TIMEOUT);
  test('synthetic after proposals', async () => {
    expect(await parity(synthetic(), true)).toBe(0);
  }, TIMEOUT);
});

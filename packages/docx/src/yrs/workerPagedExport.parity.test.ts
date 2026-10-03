import { beforeAll, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { basename, join, relative, resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import type { DocxLayoutMap, DocxPageExportOptions, DocxPagedStructuredContent } from './pagedExport';
import type { DocxParagraphIdentitySnapshot } from './paragraphIdentity';
import { createProposalRegistry, type DocxProposalSession } from './proposals';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
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

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

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

/** The text of every story and the page map's pages, which source placement does not change. */
function textAndPages(reply: PagedExport): unknown {
  if (!reply.ok) return null;
  const text = (value: unknown): string => {
    if (Array.isArray(value)) return value.map(text).join('');
    if (!value || typeof value !== 'object') return '';
    const node = value as Record<string, unknown>;
    if (node.kind === 'text' && typeof node.text === 'string') return node.text;
    return Object.entries(node).filter(([key]) => key !== 'anchor').map(([, item]) => text(item)).join('');
  };
  return {
    stories: reply.content.structured.stories.map((story) => text(story)),
    pages: reply.content.layout.pages,
  };
}

/**
 * Whether `main` differs from `worker` only by source information a session that was not seeded
 * from the package lacks: comment authors and dates, or source breaks and omitted inline source
 * content, which `main` reports with a `provenance-unavailable` diagnostic and which leave every
 * story's text and the pages unchanged.
 */
function sourceProvenanceGap(worker: PagedExport, main: PagedExport): boolean {
  if (!worker.ok || !main.ok) return false;
  const reported = (reply: typeof main) => reply.content.structured.diagnostics.some(
    ({ code }) => code === 'provenance-unavailable'
  );
  if (reported(worker)) return false;
  if (reported(main)) {
    return JSON.stringify(textAndPages(worker)) === JSON.stringify(textAndPages(main));
  }
  return leaves(worker, main).every(([path, ours, theirs]) =>
    path === '$.content.layout.exportFingerprint' ||
    (/^\$\.content\.structured\.stories\.\d+\.comment\.(author|date)$/.test(path) &&
      theirs === null && typeof ours === 'string')
  );
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
  session: DocxProposalSession,
  identities: DocxParagraphIdentitySnapshot,
  expectVersion: string
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
  const registry = createProposalRegistry(session);
  try {
    const result = registry.propose({
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
  } finally {
    registry.destroy();
  }
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
      propose(worker.proposalEngine, worker.paragraphIdentities(), exported.version);
      propose(seeded, seeded.paragraphIdentities(), seeded.version());
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

function synthetic(): Uint8Array {
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
  xml('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${paragraph('First paragraph')}<w:p>${comments}${revisions}${image}</w:p>${table}<w:p><w:pPr><w:sectPr>${section(1)}</w:sectPr></w:pPr><w:r><w:t>Section boundary</w:t></w:r></w:p>${paragraph('Second section')}<w:sectPr>${section(2)}</w:sectPr></w:body></w:document>`);
  parts.set('word/media/pixel.png', new Uint8Array(Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aD1sAAAAASUVORK5CYII=',
    'base64'
  )));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

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
    expect(await parity(synthetic())).toBeGreaterThan(0);
  }, TIMEOUT);
  test('synthetic after proposals', async () => {
    expect(await parity(synthetic(), true)).toBeGreaterThan(0);
  }, TIMEOUT);
});

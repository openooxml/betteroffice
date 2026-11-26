import { expect } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { assertSafeTree, decodeS9Blocks, decodeS9EnvelopeValue } from '../docx/rustParseFacade';
import type { BlockContent, Document } from '../types/document';
import { preloadEditWasm, type EditSession } from '../wasm/edit';
import { dirtyProjectionStory } from './dirtyProjectionStories';
import type { YrsSession } from './index';
import { projectionAllowlist, type ProjectionFamily } from './projectionParity.allowlist';
import { headerFooterStory } from './sessionInternals';
import { blockSdtStoryId, tableCellStoryId } from './storyAttributes';
import { seededCellIndices } from './tableRowSpans';
import { projectedSessionKey, stableStringify } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE_ROOT = resolve(import.meta.dir, '__fixtures__/seed-parity');

const AUTHOR = { name: 'Differential', date: '2026-01-01T00:00:00Z' };

export function preloadSeedParityWasm(): Promise<void> {
  return preloadEditWasm(new Uint8Array(readFileSync(WASM)));
}

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DEFAULT_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240" w:after="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
</w:styles>`;

const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const CONTENT_TYPES: Record<string, string> = {
  'word/document.xml': `${OFFICE_DOC}.wordprocessingml.document.main+xml`,
  'word/styles.xml': `${OFFICE_DOC}.wordprocessingml.styles+xml`,
  'word/numbering.xml': `${OFFICE_DOC}.wordprocessingml.numbering+xml`,
  'word/settings.xml': `${OFFICE_DOC}.wordprocessingml.settings+xml`,
  'word/footnotes.xml': `${OFFICE_DOC}.wordprocessingml.footnotes+xml`,
  'word/endnotes.xml': `${OFFICE_DOC}.wordprocessingml.endnotes+xml`,
  'word/comments.xml': `${OFFICE_DOC}.wordprocessingml.comments+xml`,
  'word/commentsExtended.xml': `${OFFICE_DOC}.wordprocessingml.commentsExtended+xml`,
  'word/commentsIds.xml': `${OFFICE_DOC}.wordprocessingml.commentsIds+xml`,
  'word/fontTable.xml': `${OFFICE_DOC}.wordprocessingml.fontTable+xml`,
  'word/theme/theme1.xml': `${OFFICE_DOC}.theme+xml`,
};

function partContentType(path: string): string | undefined {
  if (CONTENT_TYPES[path]) return CONTENT_TYPES[path];
  if (/^word\/header\d+\.xml$/.test(path)) return `${OFFICE_DOC}.wordprocessingml.header+xml`;
  if (/^word\/footer\d+\.xml$/.test(path)) return `${OFFICE_DOC}.wordprocessingml.footer+xml`;
  return undefined;
}

function contentTypesXml(paths: string[]): string {
  const overrides = paths
    .map((path) => {
      const type = partContentType(path);
      return type ? `  <Override PartName="/${path}" ContentType="${type}"/>` : '';
    })
    .filter(Boolean)
    .join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
${overrides}
</Types>`;
}

function walk(dir: string, prefix = ''): string[] {
  return readdirSync(dir, { withFileTypes: true })
    .flatMap((entry) =>
      entry.isDirectory()
        ? walk(join(dir, entry.name), `${prefix}${entry.name}/`)
        : [`${prefix}${entry.name}`]
    )
    .sort();
}

/**
 * Assemble a fixture directory into DOCX bytes. Every file is a package part and a
 * `<path>.base64` file decodes to the binary part at `<path>`; the content types, package
 * rels and styles are supplied when the fixture does not author its own.
 */
export function buildFixtureDocx(name: string): Uint8Array {
  const dir = join(FIXTURE_ROOT, name);
  const authored = new Map<string, Uint8Array>();
  for (const part of walk(dir)) {
    const bytes = readFileSync(join(dir, part));
    if (part.endsWith('.base64')) {
      const decoded = Buffer.from(bytes.toString('utf8').replace(/\s+/g, ''), 'base64');
      authored.set(part.slice(0, -'.base64'.length), new Uint8Array(decoded));
    } else {
      authored.set(part, new Uint8Array(bytes));
    }
  }
  return buildDocxFromParts(authored);
}

export function buildDocxFromParts(authored: PartsMap): Uint8Array {
  if (!authored.has('word/styles.xml')) authored.set('word/styles.xml', toBytes(DEFAULT_STYLES));
  if (!authored.has('_rels/.rels')) authored.set('_rels/.rels', toBytes(PACKAGE_RELS));

  const pkg: PartsMap = new Map();
  pkg.set(
    '[Content_Types].xml',
    authored.get('[Content_Types].xml') ?? toBytes(contentTypesXml([...authored.keys()]))
  );
  for (const [path, bytes] of authored) {
    if (path === '[Content_Types].xml') continue;
    pkg.set(path, bytes);
  }
  return new Uint8Array(rezipPartsToArrayBuffer(pkg));
}

export function seedParityFixtures(): string[] {
  return readdirSync(FIXTURE_ROOT, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();
}

export function random(seed: string): () => number {
  let state = 2166136261;
  for (const character of seed) state = Math.imul(state ^ character.charCodeAt(0), 16777619);
  return () => {
    state = (state + 0x6d2b79f5) | 0;
    let value = Math.imul(state ^ (state >>> 15), 1 | state);
    value = (value + Math.imul(value ^ (value >>> 7), 61 | value)) ^ value;
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

export interface ProjectionEdit {
  kind: string;
  story: string;
  paraId: string;
  refusal?: string;
}

export function edit(
  session: YrsSession,
  next: () => number,
  record?: (outcome: ProjectionEdit) => void
): string[] {
  const pick = <T>(items: readonly T[]) => items[Math.floor(next() * items.length)]!;
  const candidates = session
    .storyIds()
    .flatMap((story) =>
      session
        .paragraphs(story)
        .filter((paragraph) => paragraph.text.length > 1)
        .map((paragraph) => ({ story, paragraph }))
    );
  if (candidates.length === 0) return [];
  const body = candidates.filter(({ story }) => story === 'body');
  const pool = body.length > 0 && next() < 0.7 ? body : candidates;
  const structural = next() < 0.15;
  const count = structural || next() < 0.5 ? 1 : 2 + Math.floor(next() * 4);
  const targets = new Map<string, (typeof candidates)[number]>();
  for (let attempt = 0; attempt < count * 4 && targets.size < count; attempt += 1) {
    const target = pick(pool);
    targets.set(`${target.story}\u0000${target.paragraph.paraId}`, target);
  }
  const applied: string[] = [];
  for (const { story, paragraph } of targets.values()) {
    const length = paragraph.text.length;
    const offset = Math.floor(next() * length);
    const end = Math.min(length, offset + 1 + Math.floor(next() * 4));
    const range = {
      story,
      start: { paraId: paragraph.paraId, offset },
      end: { paraId: paragraph.paraId, offset: end },
    };
    const kind = structural
      ? pick(['split', 'merge'] as const)
      : pick(['insert', 'delete', 'replace', 'bold', 'suggestInsert', 'suggestDelete'] as const);
    const version = record ? session.version() : undefined;
    const outcome: ProjectionEdit = { kind, story, paraId: paragraph.paraId };
    try {
      if (kind === 'insert') session.insertText({ story, paraId: paragraph.paraId, offset }, 'Q9Z');
      else if (kind === 'delete') session.deleteRange(range);
      else if (kind === 'replace') session.replaceRange(range, 'Q9Z');
      else if (kind === 'bold') session.toggleMark(range, { type: 'bold' });
      else if (kind === 'suggestInsert')
        session.insertText({ story, paraId: paragraph.paraId, offset }, 'Q9Z', AUTHOR);
      else if (kind === 'suggestDelete') session.deleteRange(range, AUTHOR);
      else if (kind === 'split') session.splitParagraph({ story, paraId: paragraph.paraId, offset });
      else session.mergeParagraphs(story, paragraph.paraId);
      applied.push(`${kind}@${story}`);
    } catch (error) {
      record?.({ ...outcome, refusal: String(error) });
      continue;
    }
    if (record) {
      record(session.version() === version ? { ...outcome, refusal: 'no state change' } : outcome);
    }
  }
  return applied;
}

export function projectionRootOrder(storyIds: Iterable<string>): string[] {
  return [...storyIds].sort();
}

function normalizeProjection(value: unknown, semantic: boolean): unknown {
  if (value instanceof ArrayBuffer) return { base64: Buffer.from(value).toString('base64') };
  if (ArrayBuffer.isView(value)) {
    return { base64: Buffer.from(value.buffer, value.byteOffset, value.byteLength).toString('base64') };
  }
  if (value instanceof Date) return value.toISOString();
  if (value instanceof Map) {
    return Object.fromEntries([...value].map(([key, entry]) => [key, normalizeProjection(entry, semantic)]));
  }
  if (Array.isArray(value)) return value.map((entry) => normalizeProjection(entry, semantic));
  const object = record(value);
  if (!object) return value;
  const result = Object.fromEntries(Object.entries(object).filter(([, entry]) => entry !== undefined)
    .map(([key, entry]) => [key, normalizeProjection(entry, semantic)]));
  if (!semantic) return result;

  // A paragraph has the same semantic value with or without its independently checked source ordinal.
  if (result.type === 'paragraph' && Array.isArray(result.content)) delete result.sourceOrdinal;

  // An absent whitespace hint and true are equivalent when the text itself forces preservation.
  if (result.type === 'text' && typeof result.text === 'string' && result.preserveSpace === true &&
      (result.text.startsWith(' ') || result.text.endsWith(' ') || result.text.includes('  '))) {
    delete result.preserveSpace;
  }

  // A run with an empty formatting object is equivalent to a run with absent formatting.
  const formatting = record(result.formatting);
  if (result.type === 'run' && formatting && Object.keys(formatting).length === 0) {
    delete result.formatting;
  }

  // An empty table style cascade is equivalent to absent cascade metadata.
  const cascade = record(formatting?.styleCascade);
  if (result.type === 'table' && formatting && cascade && Object.keys(cascade).length === 0) {
    delete formatting.styleCascade;
    if (Object.keys(formatting).length === 0) delete result.formatting;
  }

  // An omitted field tree display mode is equivalent to an explicit result mode.
  const tree = record(result.fieldTree);
  if ((result.type === 'simpleField' || result.type === 'complexField') && tree &&
      tree.displayMode === 'result') delete tree.displayMode;

  // A missing package separator collection is equivalent to an empty collection of that kind.
  const pkg = record(result.package);
  if (pkg && record(pkg.document)) {
    for (const name of ['footnoteSeparators', 'endnoteSeparators']) {
      const separators = pkg[name];
      if (Array.isArray(separators) && separators.length === 0) delete pkg[name];
    }
  }
  return result;
}

export function canonicalProjection(value: unknown): unknown {
  return normalizeProjection(value, true);
}

type ProjectionExport = 'project_document_json' | 'project_story_json';
type ProjectionSession = EditSession & Partial<{
  project_document_json(input: string): string;
  project_story_json(storyId: string, input: string): string;
}>;

interface Site {
  storyId: string;
  paragraph: number | null;
}

interface Feature {
  family: ProjectionFamily;
  kind: string;
}

export type Deferral = Site & Feature;

interface CreatedListSite {
  paragraph: number;
  slot: 'current' | 'previous';
  key: string;
}

interface ProjectionSideTables {
  paragraphSessionKeys: string[];
  sessionRevisionPlaceholders: number[];
  createdListSites: CreatedListSite[];
  deferred: Deferral[];
}

interface ProjectionWire extends ProjectionSideTables {
  document: unknown;
}

export interface NativeProjection extends ProjectionSideTables {
  document: Document;
}

interface StoryProjectionWire extends ProjectionSideTables {
  storyId: string;
  blocks: unknown;
}

interface NativeStoryProjection extends ProjectionSideTables {
  storyId: string;
  blocks: BlockContent[];
}

interface Story {
  content: readonly BlockContent[];
  blocks: readonly unknown[];
  keys: readonly (string | undefined)[];
  owner?: Site & Feature;
}

export type Allowance = (typeof projectionAllowlist)[number];
export type AllowanceUsage = Map<Allowance, number>;
const families = new Set<ProjectionFamily>(['fields', 'tables', 'markers', 'storiesAux']);

export function projectionExport<K extends ProjectionExport>(session: EditSession, name: K) {
  const method = (session as ProjectionSession)[name];
  if (typeof method !== 'function') {
    throw new Error(`DOCX projection tests require the edit WASM export EditSession.${name}`);
  }
  return method as NonNullable<ProjectionSession[K]>;
}

function validateSideTables(wire: ProjectionSideTables): void {
  expect(wire.paragraphSessionKeys.every((key) => typeof key === 'string')).toBe(true);
  expect(wire.sessionRevisionPlaceholders.every(Number.isFinite)).toBe(true);
  expect(wire.createdListSites.every(({ paragraph, slot, key }) =>
    Number.isSafeInteger(paragraph) && paragraph >= 0 &&
    (slot === 'current' || slot === 'previous') && typeof key === 'string'
  )).toBe(true);
  expect(wire.deferred.every(({ family, kind, storyId, paragraph }) =>
    families.has(family) && typeof kind === 'string' && typeof storyId === 'string' &&
    (paragraph === null || (Number.isSafeInteger(paragraph) && paragraph >= 0))
  )).toBe(true);
}

export function nativeProjection(
  raw: EditSession,
  buffer: ArrayBuffer,
  roots?: ReadonlySet<string>
): NativeProjection {
  const input = JSON.stringify(roots ? { rootStories: projectionRootOrder(roots) } : {});
  const wire = JSON.parse(projectionExport(raw, 'project_document_json').call(raw, input)) as ProjectionWire;
  validateSideTables(wire);
  const { document } = decodeS9EnvelopeValue({
    wireVersion: 1,
    document: wire.document,
    embeddedFontParts: [],
  }, buffer);
  return { ...wire, document };
}

export function nativeStoryProjection(raw: EditSession, storyId: string): NativeStoryProjection {
  const json = projectionExport(raw, 'project_story_json').call(raw, storyId, '{}');
  const wire = JSON.parse(json) as StoryProjectionWire;
  assertSafeTree(wire, 'wire');
  expect(Object.keys(wire).sort()).toEqual([
    'blocks', 'createdListSites', 'deferred', 'paragraphSessionKeys', 'sessionRevisionPlaceholders', 'storyId',
  ]);
  expect(wire.storyId).toBe(storyId);
  validateSideTables(wire);
  return { ...wire, blocks: decodeS9Blocks(wire.blocks, 'wire.blocks') };
}

function record(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined;
}

function equal(expected: unknown, actual: unknown): boolean {
  return expected === undefined || actual === undefined
    ? expected === actual
    : stableStringify(expected) === stableStringify(actual);
}

function typeOf(value: unknown): string | undefined {
  const type = record(value)?.type;
  return typeof type === 'string' ? type : undefined;
}

function collectStory(
  stories: Map<string, Story>,
  storyId: string,
  blocks: readonly BlockContent[],
  owner?: Story['owner']
): void {
  if (stories.has(storyId)) return;
  const keys: Story['keys'][number][] = [];
  const skeleton: unknown[] = [];
  stories.set(storyId, { content: blocks, blocks: skeleton, keys, owner });
  let tableIndex = 0;
  let sdtIndex = 0;
  for (const block of blocks) {
    const site = { storyId, paragraph: keys.length, kind: block.type };
    if (block.type === 'paragraph') keys.push(projectedSessionKey(block));
    if (block.type === 'blockSdt') {
      const child = blockSdtStoryId(storyId, sdtIndex++);
      collectStory(stories, child, block.content, { ...site, family: 'storiesAux' });
      skeleton.push(canonicalProjection({ ...block, content: { storyId: child } }));
    } else if (block.type === 'table') {
      const table = tableIndex++;
      const cells = seededCellIndices(block);
      const rows = block.rows.map((row, rowIndex) => ({
        ...row,
        cells: row.cells.map((cell, cellIndex) => {
          const ordinal = cells[rowIndex]!.indexOf(cellIndex);
          if (ordinal < 0) return cell;
          const child = tableCellStoryId(storyId, table, rowIndex, ordinal);
          collectStory(stories, child, cell.content, { ...site, family: 'tables' });
          return { ...cell, content: { storyId: child } };
        }),
      }));
      skeleton.push(canonicalProjection({ ...block, rows }));
    } else {
      skeleton.push(canonicalProjection(block));
    }
  }
}

export function canonicalStories(session: YrsSession, document: Document): Map<string, Story> {
  const stories = new Map<string, Story>();
  const pkg = document.package;
  collectStory(stories, 'body', pkg.document.content);
  const furniture = [...(pkg.headers ?? []), ...(pkg.footers ?? [])];
  for (const [id, part] of furniture) {
    const canonical = headerFooterStory(session, id);
    const canonicalPart = furniture.find(([relId]) => canonical === `hf:${relId}`)?.[1] ?? part;
    collectStory(stories, canonical, canonicalPart.content, { storyId: canonical, paragraph: null, family: 'storiesAux', kind: 'headerFooterStory' });
  }
  for (const [prefix, notes] of [['fn', pkg.footnotes], ['en', pkg.endnotes]] as const) {
    for (const note of notes ?? []) {
      const storyId = `${prefix}:${note.id}`;
      collectStory(stories, storyId, note.content, { storyId, paragraph: null, family: 'storiesAux', kind: 'noteStory' });
    }
  }
  return stories;
}

function packageFields(session: YrsSession, document: Document): unknown {
  const { package: pkg, ...envelope } = document;
  const { document: body, headers, footers, footnotes, endnotes, ...fields } = pkg;
  const parts = (values: typeof headers) => values && new Map(
    [...values].map(([id, { content, ...part }]) => {
      const canonical = headerFooterStory(session, id).slice(3);
      if (canonical === id) return [id, part];
      const owner = values.get(canonical);
      const sameStory = owner && equal(canonicalProjection(content), canonicalProjection(owner.content));
      return [id, sameStory ? part : { ...part, content }];
    })
  );
  let nextSectionStart = 0;
  const notes = (values: typeof footnotes | typeof endnotes) =>
    values?.map(({ content: _, ...note }) => note);
  return canonicalProjection({
    ...envelope,
    package: {
      ...fields,
      headers: parts(headers),
      footers: parts(footers),
      footnotes: notes(footnotes),
      endnotes: notes(endnotes),
      document: {
        ...body,
        content: undefined,
        sections: body.sections?.map((section) => {
          const first = section.content[0];
          const bodyIndex = first === undefined ? -1 : body.content.indexOf(first);
          const start = bodyIndex === -1 ? nextSectionStart : bodyIndex;
          const end = start + section.content.length;
          nextSectionStart = end;
          return { ...section, content: [start, end] };
        }),
      },
    },
  });
}

const DEFERRED_NODE_KINDS: Readonly<Record<string, string>> = {
  simpleField: 'field',
  complexField: 'field',
  fieldChar: 'field',
  instrText: 'field',
  hyperlink: 'hyperlink',
  tab: 'tab',
  break: 'break',
  footnoteRef: 'noteRef',
  endnoteRef: 'noteRef',
  inlineSdt: 'sdt',
  mathEquation: 'math',
  drawing: 'image',
  chart: 'chart',
  opaqueDrawing: 'opaqueDrawing',
  shape: 'shape',
};

const MARKER_KINDS: Readonly<Record<string, string>> = {
  rawXml: 'rawInline',
  commentReference: 'field',
  commentRangeStart: 'commentRanges',
  commentRangeEnd: 'commentRanges',
  bookmarkStart: 'bookmarks',
  bookmarkEnd: 'bookmarks',
  insertion: 'ins',
  moveTo: 'ins',
  deletion: 'del',
  moveFrom: 'del',
};

function deferredNode(value: unknown): Feature | undefined {
  const type = typeOf(value);
  if (type === undefined) return undefined;
  if (type === 'table') return { family: 'tables', kind: type };
  if (type === 'blockSdt') return { family: 'storiesAux', kind: type };
  const marker = MARKER_KINDS[type];
  if (marker) return { family: 'markers', kind: marker };
  const breakType = type === 'break' ? record(value)?.breakType : undefined;
  if (breakType === 'page' || breakType === 'column') {
    return { family: 'fields', kind: `${breakType}Break` };
  }
  const field = DEFERRED_NODE_KINDS[type];
  return field ? { family: 'fields', kind: field } : undefined;
}

function wrapperChildren(node: unknown, kind: string): unknown[] {
  const value = record(node);
  const children = kind === 'hyperlink'
    ? value?.structuredChildren ?? value?.children
    : kind === 'ins' || kind === 'del' ? value?.content : undefined;
  return Array.isArray(children) ? children : [];
}

function arraySite(site: Site, path: string, paragraph: number, value: unknown): Site {
  if (path === 'blocks') return { ...site, paragraph };
  if (path === 'package.document.sections') return { storyId: 'body', paragraph: null };
  const id = record(value)?.id;
  if (typeof id === 'number') {
    if (path === 'package.footnotes') return { storyId: `fn:${id}`, paragraph: null };
    if (path === 'package.endnotes') return { storyId: `en:${id}`, paragraph: null };
  }
  return site;
}

function deferralParagraph(feature: Feature, site: Site): number | null {
  return feature.family === 'markers' && (feature.kind === 'commentRanges' || feature.kind === 'revisions')
    ? null : site.paragraph;
}

function plainTextRun(value: unknown): { text: string; attributes: unknown } | undefined {
  const run = record(value);
  if (run?.type !== 'run' || !Array.isArray(run.content) || run.content.length !== 1) return undefined;
  const child = record(run.content[0]);
  if (child?.type !== 'text' || typeof child.text !== 'string') return undefined;
  const { content: _, ...attributes } = run;
  const { text, ...textAttributes } = child;
  return { text, attributes: [attributes, textAttributes] };
}

interface ArrayEntry {
  value: unknown;
  index: number;
}

class Comparison {
  readonly differences: string[] = [];

  constructor(
    private readonly deferred: readonly Deferral[] = [],
    private readonly strict = false,
    private readonly used?: AllowanceUsage
  ) {}

  allow(feature: Feature, path: string, site: Site, count = 1): boolean {
    const match = this.allowance(feature, path, site);
    if (!match) return false;
    this.recordAllowance(match, count);
    return true;
  }

  private recordAllowance(allowance: Allowance, count = 1): void {
    if (this.used) this.used.set(allowance, (this.used.get(allowance) ?? 0) + count);
  }

  private allowance({ family, kind }: Feature, path: string, site: Site): Allowance | undefined {
    if (this.strict) return undefined;
    const shape = path.replace(/\[\d+\]/g, '[]');
    const paragraph = deferralParagraph({ family, kind }, site);
    const reported = this.deferred.some((deferral) =>
      deferral.family === family && deferral.kind === kind &&
      deferral.storyId === site.storyId && deferral.paragraph === paragraph
    );
    return reported ? projectionAllowlist.find((entry) =>
      entry.family === family && entry.path === shape && entry.kind === kind
    ) : undefined;
  }

  report(site: Site, path: string, expected: unknown, actual: unknown): void {
    const display = (value: unknown) => value === undefined ? '<missing>' : stableStringify(value).slice(0, 200);
    this.differences.push(
      `story ${site.storyId}, paragraph ${site.paragraph ?? '-'}, ${path}: TS ${display(expected)}; native ${display(actual)}`
    );
  }

  compare(expected: unknown, actual: unknown, site: Site, path: string, feature?: Feature): void {
    if (equal(expected, actual)) return;
    if (Array.isArray(expected) && Array.isArray(actual)) {
      this.compareArray(expected, actual, site, path, feature);
      return;
    }
    const left = record(expected);
    const right = record(actual);
    if (left && right) {
      const owner = typeOf(left) === typeOf(right) ? deferredNode(right) ?? feature : feature;
      for (const field of new Set([...Object.keys(left), ...Object.keys(right)])) {
        const childPath = path ? `${path}.${field}` : field;
        const missing = this.missingField(left[field], right[field], childPath, owner);
        if (missing && this.allow(missing, childPath, site)) continue;
        this.compare(left[field], right[field], site, childPath, owner);
      }
      return;
    }
    this.report(site, path, expected, actual);
  }

  private missingField(expected: unknown, actual: unknown, path: string, owner?: Feature): Feature | undefined {
    if (expected !== undefined && actual === undefined) {
      if (/^blocks\[\d+\]\.content\[\d+\]\.propertyChanges$/.test(path)) {
        return { family: 'markers', kind: 'rPrChange' };
      }
    }
    if (!equal(expected, actual) && /^package\.document\.sections\[\d+\]\.(content|id|properties\.\w+)$/.test(path)) {
      return { family: 'storiesAux', kind: 'sectionRanges' };
    }
    if (owner?.family === 'tables' && expected !== undefined && actual === undefined) return owner;
    if (owner?.family === 'fields' && expected === undefined && actual !== undefined) return owner;
    return undefined;
  }

  private compareArray(expected: unknown[], actual: unknown[], site: Site, path: string, owner?: Feature): void {
    let right = 0;
    let paragraph = 0;
    const pending = expected.map((value, index) => ({ value, index }));
    const native = [...actual];
    for (let left = 0; left < pending.length; left++) {
      const { value, index } = pending[left]!;
      const location = arraySite(site, path, paragraph, value);
      const entryPath = `${path}[${index}]`;
      const candidate = native[right];
      const remaining = native.slice(right);
      const omission = this.omission(value, remaining, location, entryPath);
      if (omission) {
        this.recordAllowance(omission.allowance);
        pending.splice(left + 1, 0, ...omission.children.map((child) => ({ value: child, index })));
        continue;
      }
      const joinedUntil = this.joinAcrossOmissions(pending, left, remaining, location, path);
      if (joinedUntil !== undefined) {
        left = joinedUntil;
        right++;
        continue;
      }
      const retained = this.retainedInline(candidate, pending.slice(left).map(({ value }) => value), location, entryPath);
      if (retained) {
        this.recordAllowance(retained.allowance);
        native.splice(right, 1, ...retained.children);
        left--;
        continue;
      }
      this.compare(value, candidate, location, entryPath, owner);
      if (right < native.length) right++;
      if (path === 'blocks' && typeOf(value) === 'paragraph') paragraph++;
    }
    for (; right < native.length; right++) {
      const value = native[right];
      const location = arraySite(site, path, paragraph, value);
      const entryPath = `${path}[${right}]`;
      const retained = this.retainedInline(value, [], location, entryPath);
      if (retained && retained.children.length === 0) this.recordAllowance(retained.allowance);
      else if (!owner || !this.allow(owner, entryPath, location)) this.report(location, entryPath, undefined, value);
      if (path === 'blocks' && typeOf(value) === 'paragraph') paragraph++;
    }
  }

  private retainedInline(value: unknown, expected: readonly unknown[], site: Site, path: string) {
    if (value === undefined || !/^blocks\[\d+\]\.content\[\d+\]$/.test(path)) return undefined;
    const omission = this.omission(value, expected, site, path);
    const family = omission?.allowance.family;
    return family === 'fields' || family === 'markers' ? omission : undefined;
  }

  private omission(value: unknown, actual: readonly unknown[], site: Site, path: string) {
    if (this.strict) return undefined;
    const candidate = actual[0];
    const feature = deferredNode(value);
    if (feature && typeOf(value) !== typeOf(candidate)) {
      const match = this.allowance(feature, path, site);
      if (match && !actual.some((node) => equal(value, node))) {
        return { allowance: match, children: wrapperChildren(value, feature.kind) };
      }
    }
    const content = typeOf(value) === 'run' ? record(value)?.content : undefined;
    if (!Array.isArray(content) || content.length !== 1) return undefined;
    const token = deferredNode(content[0]);
    if (!token) return undefined;
    const match = this.allowance(token, `${path}.content[0]`, site);
    const candidateContent = typeOf(candidate) === 'run' ? record(candidate)?.content : undefined;
    if (!match || (Array.isArray(candidateContent) &&
        candidateContent.some((node) => equal(token, deferredNode(node)))) ||
        actual.some((node) => equal(value, node))) return undefined;
    return { allowance: match, children: [] };
  }

  private joinAcrossOmissions(
    expected: readonly ArrayEntry[],
    start: number,
    actual: readonly unknown[],
    site: Site,
    path: string
  ): number | undefined {
    if (this.strict || !/^blocks\[\d+\]\.content$/.test(path)) return undefined;
    const first = plainTextRun(expected[start]?.value);
    const target = plainTextRun(actual[0]);
    if (!first || !target || first.text === target.text ||
        !target.text.startsWith(first.text) || !equal(first.attributes, target.attributes)) return undefined;
    const matches: Allowance[] = [];
    let text = first.text;
    for (let index = start + 1; index < expected.length; index++) {
      const entry = expected[index]!;
      const omission = this.omission(entry.value, actual, site, `${path}[${entry.index}]`);
      const adjacent = matches.length > 0;
      if (omission) matches.push(omission.allowance);
      const nodes = omission ? omission.children : [entry.value];
      if (!adjacent && !omission) return undefined;
      for (const node of nodes) {
        const run = plainTextRun(node);
        if (!run || !equal(run.attributes, target.attributes)) return undefined;
        text += run.text;
        if (!target.text.startsWith(text)) return undefined;
      }
      if (text === target.text) {
        for (const match of matches) this.recordAllowance(match);
        return index;
      }
      if (!omission) return undefined;
    }
    return undefined;
  }
}

function compareStories(
  comparison: Comparison,
  expected: Map<string, Story>,
  actual: Map<string, Story>,
  source: Map<string, Story>,
  only?: ReadonlySet<string>
): void {
  for (const storyId of new Set([...expected.keys(), ...actual.keys()])) {
    if (only && !only.has(dirtyProjectionStory(storyId))) continue;
    const left = expected.get(storyId);
    const right = actual.get(storyId);
    const site = { storyId, paragraph: null };
    if (left?.owner && !equal(left.blocks, right?.blocks) &&
        (!right || equal(right.blocks, source.get(storyId)?.blocks))) {
      const retained = new Comparison([], true);
      retained.compare(left.blocks, right?.blocks, site, 'blocks');
      let owner: Story['owner'] = left.owner;
      let allowed = false;
      while (owner) {
        const path = owner.paragraph === null ? 'blocks' : 'blocks[]';
        if ((right || owner.paragraph !== null) &&
            comparison.allow(owner, path, owner, retained.differences.length)) {
          allowed = true;
          break;
        }
        const parent: Story['owner'] = expected.get(owner.storyId)?.owner;
        if (parent === owner) break;
        owner = parent;
      }
      if (allowed) continue;
    }
    comparison.compare(left?.blocks, right?.blocks, site, 'blocks');
  }
}

function compareStoryKeys(
  comparison: Comparison,
  storyId: string,
  story: Story | undefined,
  actual: readonly string[]
): void {
  const expected = story?.keys.flatMap((key, paragraph) =>
    key === undefined ? [] : [{ key, paragraph }]
  ) ?? [];
  for (let index = 0; index < Math.max(expected.length, actual.length); index++) {
    comparison.compare(expected[index]?.key, actual[index], {
      storyId,
      paragraph: expected[index]?.paragraph ?? index,
    }, `paragraphSessionKeys[${index}]`);
  }
}

export function checkSourceOrdinals(
  storyId: string,
  blocks: readonly BlockContent[],
  source: readonly BlockContent[],
  keys: readonly string[] = []
): string[] {
  const comparison = new Comparison([], true);
  const paragraphs = (content: readonly BlockContent[]) => content.filter((block) => block.type === 'paragraph');
  const originals = paragraphs(source);
  const byKey = new Map(originals.map((paragraph, index) => [
    (!paragraph.repeatedParaId && paragraph.paraId) || `${storyId}:p${index}`, paragraph,
  ]));
  paragraphs(blocks).forEach((paragraph, index) => {
    const key = keys[index];
    const original = key === undefined ? originals[index] : byKey.get(key);
    comparison.compare(
      original?.sourceOrdinal, paragraph.sourceOrdinal, { storyId, paragraph: index }, 'sourceOrdinal'
    );
  });

  const nested = (content: readonly BlockContent[], visit: (paragraph: Record<string, unknown>, path: string) => void) => {
    const walk = (value: unknown, path: string): void => {
      if (Array.isArray(value)) {
        value.forEach((entry, index) => walk(entry, `${path}[${index}]`));
        return;
      }
      const node = record(value);
      if (!node) return;
      if (node.type === 'paragraph' && Array.isArray(node.content)) visit(node, path);
      for (const [name, child] of Object.entries(node)) {
        if (typeof child === 'object') walk(child, `${path}.${name}`);
      }
    };
    content.forEach((block, index) => {
      const path = `blocks[${index}]`;
      if (block.type === 'paragraph') walk(block.content, `${path}.content`);
      else walk(block, path);
    });
  };
  const nestedSources = new Map<string, unknown[]>();
  nested(source, (paragraph) => {
    const key = stableStringify(canonicalProjection(paragraph));
    const ordinals = nestedSources.get(key) ?? [];
    ordinals.push(paragraph.sourceOrdinal);
    nestedSources.set(key, ordinals);
  });
  nested(blocks, (paragraph, path) => {
    const key = stableStringify(canonicalProjection(paragraph));
    const candidates = nestedSources.get(key);
    const match = candidates?.indexOf(paragraph.sourceOrdinal) ?? -1;
    if (match >= 0) candidates!.splice(0, match + 1);
    else comparison.compare(candidates?.[0], paragraph.sourceOrdinal, { storyId, paragraph: null }, `${path}.sourceOrdinal`);
  });
  return comparison.differences;
}

export function check(
  session: YrsSession,
  ts: Document,
  native: NativeProjection,
  source: Map<string, Story>,
  roots?: ReadonlySet<string>,
  used?: AllowanceUsage
): string[] {
  const comparison = new Comparison(native.deferred, false, used);
  const expected = canonicalStories(session, ts);
  const actual = canonicalStories(session, native.document);
  compareStories(comparison, expected, actual, source);
  if (!roots || roots.has('body')) {
    compareStoryKeys(comparison, 'body', expected.get('body'), native.paragraphSessionKeys);
  }
  for (const [storyId, story] of actual) {
    if (dirtyProjectionStory(storyId) !== storyId) continue;
    comparison.differences.push(...checkSourceOrdinals(
      storyId, story.content, source.get(storyId)?.content ?? [],
      storyId === 'body' ? native.paragraphSessionKeys : []
    ));
  }
  comparison.compare(packageFields(session, ts), packageFields(session, native.document), {
    storyId: '<package>', paragraph: null,
  }, '');
  return comparison.differences;
}

export function checkStory(
  expected: Map<string, Story>,
  native: NativeStoryProjection,
  source: Map<string, Story>,
  used?: AllowanceUsage
): string[] {
  const { storyId, blocks, deferred, paragraphSessionKeys } = native;
  const comparison = new Comparison(deferred, false, used);
  const actual = new Map<string, Story>();
  collectStory(actual, storyId, blocks);
  compareStories(comparison, expected, actual, source, new Set([storyId]));
  const rootDeferred = deferred.some((entry) => entry.storyId === storyId && entry.paragraph === null &&
    entry.family === 'storiesAux' && (entry.kind === 'headerFooterStory' || entry.kind === 'noteStory'));
  if (!rootDeferred) compareStoryKeys(comparison, storyId, expected.get(storyId), paragraphSessionKeys);
  comparison.differences.push(...checkSourceOrdinals(storyId, blocks, source.get(storyId)?.content ?? [], paragraphSessionKeys));
  return comparison.differences;
}

export function checkComposition(
  session: YrsSession,
  stories: NativeStoryProjection[],
  native: NativeProjection
): string[] {
  const comparison = new Comparison([], true);
  const document = canonicalStories(session, native.document);
  const keys: string[] = [];
  const placeholders = new Set<number>();
  const sites: CreatedListSite[] = [];
  const deferred: Deferral[] = [];
  for (const story of stories) {
    comparison.compare(
      normalizeProjection(document.get(story.storyId)?.content, false),
      normalizeProjection(story.blocks, false),
      { storyId: story.storyId, paragraph: null },
      'blocks'
    );
    const offset = keys.length;
    keys.push(...story.paragraphSessionKeys);
    for (const id of story.sessionRevisionPlaceholders) placeholders.add(id);
    for (const site of story.createdListSites) sites.push({ ...site, paragraph: site.paragraph + offset });
    deferred.push(...story.deferred);
  }
  const location = { storyId: '<package>', paragraph: null };
  comparison.compare(native.paragraphSessionKeys, keys, location, 'paragraphSessionKeys');
  comparison.compare(native.sessionRevisionPlaceholders, [...placeholders].sort((a, b) => a - b), location,
    'sessionRevisionPlaceholders');
  comparison.compare(native.createdListSites, sites, location, 'createdListSites');
  comparison.compare(native.deferred, deferred, location, 'deferred');
  return comparison.differences;
}

export function compareSnapshots(
  session: YrsSession,
  expected: Document,
  actual: Document,
  roots: ReadonlySet<string> | undefined
): string[] {
  const comparison = new Comparison([], true);
  compareStories(comparison, canonicalStories(session, expected), canonicalStories(session, actual), new Map(), roots);
  if (!roots) {
    comparison.compare(packageFields(session, expected), packageFields(session, actual), { storyId: '<package>', paragraph: null }, '');
  }
  return comparison.differences;
}

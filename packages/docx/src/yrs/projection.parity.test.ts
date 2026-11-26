import { afterAll, beforeAll, describe, expect, it } from 'bun:test';

import { toBytes } from '../docx/rezip/parts';
import { assertSafeTree, decodeS9Blocks, decodeS9EnvelopeValue } from '../docx/rustParseFacade';
import type { BlockContent, Document } from '../types/document';
import { createEditSession, type EditSession } from '../wasm/edit';
import { dirtyProjectionStory } from './dirtyProjectionStories';
import type { YrsSession } from './index';
import { projectionAllowlist, type ProjectionFamily } from './projectionParity.allowlist';
import {
  buildDocxFromParts,
  buildFixtureDocx,
  edit,
  normalizeProjection,
  preloadSeedParityWasm,
  projectionRootOrder,
  random,
  seedParityFixtures,
  type ProjectionEdit,
} from './projectionParity.support';
import { headerFooterStory } from './sessionInternals';
import { blockSdtStoryId, tableCellStoryId } from './storyAttributes';
import { seededCellIndices } from './tableRowSpans';
import { wrapSession } from './yrsSessionFacade';
import { projectedSessionKey, stableStringify, yrsToDocument } from './yrsToDocument';

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

type Deferral = Site & Feature;

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

interface NativeProjection extends ProjectionSideTables {
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

type Allowance = (typeof projectionAllowlist)[number];
const used = new Set<Allowance>();
const families = new Set<ProjectionFamily>(['fields', 'tables', 'markers', 'storiesAux']);
const CLIENT = 48001;
const PARSER_TABLE_LEVELS = 64;
const DEEP_MARKER = 'innermost-projection-depth-marker';

function projectionExport<K extends ProjectionExport>(session: EditSession, name: K) {
  const method = (session as ProjectionSession)[name];
  if (typeof method !== 'function') {
    throw new Error(`DOCX projection parity requires the edit WASM export EditSession.${name}`);
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

function nativeProjection(
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

function nativeStoryProjection(raw: EditSession, storyId: string): NativeStoryProjection {
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
      skeleton.push(normalizeProjection({ ...block, content: { storyId: child } }));
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
      skeleton.push(normalizeProjection({ ...block, rows }));
    } else {
      skeleton.push(normalizeProjection(block));
    }
  }
}

function canonicalStories(session: YrsSession, document: Document): Map<string, Story> {
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
      const sameStory = owner && equal(normalizeProjection(content), normalizeProjection(owner.content));
      return [id, sameStory ? part : { ...part, content }];
    })
  );
  let nextSectionStart = 0;
  const notes = (values: typeof footnotes | typeof endnotes) =>
    values?.map(({ content: _, ...note }) => note);
  return normalizeProjection({
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
  hyperlink: 'hyperlink',
  tab: 'tab',
  break: 'break',
  footnoteRef: 'noteRef',
  endnoteRef: 'noteRef',
  inlineSdt: 'sdt',
  mathEquation: 'math',
  drawing: 'image',
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

  constructor(private readonly deferred: readonly Deferral[] = [], private readonly strict = false) {}

  allow(feature: Feature, path: string, site: Site): boolean {
    const match = this.allowance(feature, path, site);
    if (!match) return false;
    used.add(match);
    return true;
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

  compare(expected: unknown, actual: unknown, site: Site, path: string): void {
    if (equal(expected, actual)) return;
    if (Array.isArray(expected) && Array.isArray(actual)) {
      this.compareArray(expected, actual, site, path);
      return;
    }
    const left = record(expected);
    const right = record(actual);
    if (left && right) {
      for (const field of new Set([...Object.keys(left), ...Object.keys(right)])) {
        const childPath = path ? `${path}.${field}` : field;
        const feature = this.missingField(left[field], right[field], childPath);
        if (feature && this.allow(feature, childPath, site)) continue;
        this.compare(left[field], right[field], site, childPath);
      }
      return;
    }
    this.report(site, path, expected, actual);
  }

  private missingField(expected: unknown, actual: unknown, path: string): Feature | undefined {
    if (expected !== undefined && actual === undefined) {
      if (/^blocks\[\d+\]\.propertyChanges$/.test(path)) return { family: 'markers', kind: 'pPrChange' };
      if (/^blocks\[\d+\]\.content\[\d+\]\.propertyChanges$/.test(path)) {
        return { family: 'markers', kind: 'rPrChange' };
      }
    }
    if (expected === undefined && typeof actual === 'string' &&
        /^package\.(endnotes|footnotes)\[\d+\]\.verbatimXml$/.test(path)) {
      return { family: 'storiesAux', kind: 'noteStory' };
    }
    if (!equal(expected, actual) && /^package\.document\.sections\[\d+\]\.content$/.test(path)) {
      return { family: 'storiesAux', kind: 'sectionRanges' };
    }
    return undefined;
  }

  private compareArray(expected: unknown[], actual: unknown[], site: Site, path: string): void {
    let right = 0;
    let paragraph = 0;
    const pending = expected.map((value, index) => ({ value, index }));
    for (let left = 0; left < pending.length; left++) {
      const { value, index } = pending[left]!;
      const location = arraySite(site, path, paragraph, value);
      const entryPath = `${path}[${index}]`;
      const candidate = actual[right];
      const remaining = actual.slice(right);
      const omission = this.omission(value, remaining, location, entryPath);
      if (omission) {
        used.add(omission.allowance);
        pending.splice(left + 1, 0, ...omission.children.map((child) => ({ value: child, index })));
        continue;
      }
      const joinedUntil = this.joinAcrossOmissions(pending, left, remaining, location, path);
      if (joinedUntil !== undefined) {
        left = joinedUntil;
        right++;
        continue;
      }
      this.compare(value, candidate, location, entryPath);
      if (right < actual.length) right++;
      if (path === 'blocks' && typeOf(value) === 'paragraph') paragraph++;
    }
    for (; right < actual.length; right++) {
      const location = arraySite(site, path, paragraph, actual[right]);
      this.report(location, `${path}[${right}]`, undefined, actual[right]);
      if (path === 'blocks' && typeOf(actual[right]) === 'paragraph') paragraph++;
    }
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
        for (const match of matches) used.add(match);
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
  const omittedStories = new Set<string>();
  for (const storyId of new Set([...expected.keys(), ...actual.keys()])) {
    if (only && !only.has(dirtyProjectionStory(storyId))) continue;
    const left = expected.get(storyId);
    const right = actual.get(storyId);
    const owner = left?.owner;
    if (owner && !equal(left?.blocks, right?.blocks)) {
      if (!right && (omittedStories.has(owner.storyId) ||
          (owner.paragraph !== null && comparison.allow(owner, 'blocks[]', owner)))) {
        omittedStories.add(storyId);
        continue;
      }
      if (right && equal(right.blocks, source.get(storyId)?.blocks) &&
          comparison.allow(owner, owner.paragraph === null ? 'blocks' : 'blocks[]', owner)) continue;
    }
    comparison.compare(left?.blocks, right?.blocks, { storyId, paragraph: null }, 'blocks');
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

function check(
  session: YrsSession,
  ts: Document,
  native: NativeProjection,
  source: Map<string, Story>,
  roots?: ReadonlySet<string>
): string[] {
  const comparison = new Comparison(native.deferred);
  const expected = canonicalStories(session, ts);
  const actual = canonicalStories(session, native.document);
  compareStories(comparison, expected, actual, source);
  if (!roots || roots.has('body')) {
    compareStoryKeys(comparison, 'body', expected.get('body'), native.paragraphSessionKeys);
  }
  comparison.compare(packageFields(session, ts), packageFields(session, native.document), {
    storyId: '<package>', paragraph: null,
  }, '');
  return comparison.differences;
}

function checkStory(
  expected: Map<string, Story>,
  native: NativeStoryProjection,
  source: Map<string, Story>
): string[] {
  const { storyId, blocks, deferred, paragraphSessionKeys } = native;
  const comparison = new Comparison(deferred);
  const actual = new Map<string, Story>();
  collectStory(actual, storyId, blocks);
  compareStories(comparison, expected, actual, source, new Set([storyId]));
  const rootDeferred = deferred.some((entry) => entry.storyId === storyId && entry.paragraph === null &&
    entry.family === 'storiesAux' && (entry.kind === 'headerFooterStory' || entry.kind === 'noteStory'));
  if (!rootDeferred) compareStoryKeys(comparison, storyId, expected.get(storyId), paragraphSessionKeys);
  return comparison.differences;
}

function checkComposition(
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
      normalizeProjection(document.get(story.storyId)?.content),
      normalizeProjection(story.blocks),
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

function compareSnapshots(
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

function editParagraphProperties(session: YrsSession, recordEdit: (outcome: ProjectionEdit) => void): void {
  const paragraphs = session.paragraphs('body');
  const paragraph = paragraphs[2] ?? paragraphs[0];
  const outcome = { kind: 'paragraph properties', story: 'body', paraId: paragraph?.paraId ?? '' };
  if (!paragraph?.paraId) {
    recordEdit({ ...outcome, refusal: 'no body paragraph' });
    return;
  }
  const position = { paraId: paragraph.paraId, offset: 0 };
  const version = session.version();
  try {
    session.setParagraphAttrs({ story: 'body', start: position, end: position }, {
      lineSpacing: 360,
      lineSpacingRule: 'auto',
      indentLeft: 1080,
    });
  } catch (error) {
    recordEdit({ ...outcome, refusal: String(error) });
    return;
  }
  recordEdit(version === session.version() ? { ...outcome, refusal: 'no state change' } : outcome);
}

function syntheticDocx(body: string): Uint8Array {
  return buildDocxFromParts(new Map([['word/document.xml', toBytes(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`
  )]]));
}

function nestedTables(levels: number): string {
  const open = '<w:tbl><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc>';
  const close = '</w:tc></w:tr></w:tbl>';
  const field = '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>field result</w:t></w:r></w:fldSimple></w:p>';
  return `${field}${open.repeat(levels)}<w:p><w:r><w:t>${DEEP_MARKER}</w:t></w:r></w:p>${close.repeat(levels)}<w:p/>`;
}

function expectDeepMarker(blocks: readonly BlockContent[], levels: number): void {
  for (let depth = 0; depth < levels; depth++) {
    const table = blocks.find((block) => block.type === 'table');
    if (!table) throw new Error(`missing table level ${depth + 1}`);
    blocks = table.rows[0]!.cells[0]!.content;
  }
  expect(blocks.some((block) => block.type === 'table')).toBe(false);
  expect(JSON.stringify(blocks)).toContain(DEEP_MARKER);
}

describe('Native DOCX projection parity', () => {
  const fixtures = seedParityFixtures();

  beforeAll(async () => {
    expect(fixtures.length).toBeGreaterThan(0);
    await preloadSeedParityWasm();
    const raw = createEditSession(CLIENT);
    try {
      projectionExport(raw, 'project_document_json');
      projectionExport(raw, 'project_story_json');
    } finally {
      raw.free();
    }
  }, 30_000);

  it('carries the deepest parsed table through both native JSON exports', () => {
    const raw = createEditSession(CLIENT);
    const session = wrapSession(raw, CLIENT);
    try {
      session.seedFromDocx(syntheticDocx(nestedTables(PARSER_TABLE_LEVELS)));
      const base = session.materializeDocx()!;
      const source = canonicalStories(session, base);
      const ts = yrsToDocument(session, base);
      const native = nativeProjection(raw, base.originalBuffer!);
      const story = nativeStoryProjection(raw, 'body');
      expect(check(session, ts, native, source)).toEqual([]);
      expect(checkStory(canonicalStories(session, ts), story, source)).toEqual([]);
      expect(checkComposition(session, [story], native)).toEqual([]);
      expectDeepMarker(native.document.package.document.content, PARSER_TABLE_LEVELS);
      expectDeepMarker(story.blocks, PARSER_TABLE_LEVELS);
    } finally {
      session.destroy();
    }
  }, 30_000);

  for (const levels of [1, 260]) {
    it(`reports all ${levels} session-created tables as deferred`, () => {
      const raw = createEditSession(CLIENT);
      const session = wrapSession(raw, CLIENT);
      try {
        session.seedFromDocx(syntheticDocx('<w:p/>'));
        const base = session.materializeDocx()!;
        const source = canonicalStories(session, base);
        let storyId = 'body';
        const expected: Deferral[] = [];
        for (let depth = 0; depth < levels; depth++) {
          const { paraId } = session.paragraphs(storyId)[0]!;
          const receipt = session.insertTable({ story: storyId, paraId, offset: 0 }, 1, 1);
          expected.push({ family: 'tables', kind: 'table', storyId, paragraph: 0 });
          storyId = receipt.createdStoryIds[0]!;
        }
        const { paraId } = session.paragraphs(storyId)[0]!;
        session.insertText({ story: storyId, paraId, offset: 0 }, DEEP_MARKER);
        const ts = yrsToDocument(session, base);
        expectDeepMarker(ts.package.document.content, levels);

        const native = nativeProjection(raw, base.originalBuffer!);
        const story = nativeStoryProjection(raw, 'body');
        expect(native.deferred.filter(({ family }) => family === 'tables')).toEqual(expected);
        expect(story.deferred.filter(({ family }) => family === 'tables')).toEqual(expected);
        expect(check(session, ts, native, source)).toEqual([]);
        expect(checkStory(canonicalStories(session, ts), story, source)).toEqual([]);
        expect(checkComposition(session, [story], native)).toEqual([]);
      } finally {
        session.destroy();
      }
    }, 60_000);
  }

  for (const name of fixtures) {
    it(`matches the TypeScript projection of ${name} through deterministic edits`, () => {
      const raw = createEditSession(CLIENT);
      const session = wrapSession(raw, CLIENT);
      try {
        const bytes = buildFixtureDocx(name);
        const { document: base } = session.seedFromDocx(bytes);
        const buffer = base.originalBuffer!;
        const source = canonicalStories(session, base);
        const next = random(name);
        const failures: string[] = [];
        const checkState = (label: string, differences: string[]) => {
          if (differences.length) failures.push(`${label}\n${differences.join('\n')}`);
        };
        const checkProjections = (label: string, ts: Document, native: NativeProjection): void => {
          checkState(`${label}, document`, check(session, ts, native, source));
          const expected = canonicalStories(session, ts);
          const roots = projectionRootOrder([...expected.keys()].filter((id) => dirtyProjectionStory(id) === id));
          const stories = roots.map((id) => {
            const story = nativeStoryProjection(raw, id);
            checkState(`${label}, ${id}`, checkStory(expected, story, source));
            return story;
          });
          checkState(`${label}, composed stories versus document`, checkComposition(session, stories, native));
        };

        let ts = yrsToDocument(session, base);
        let native = nativeProjection(raw, buffer);
        checkProjections('seeded', ts, native);

        const edits: ProjectionEdit[] = [];
        for (let trial = 0; trial < 5; trial++) {
          const outcomes: ProjectionEdit[] = [];
          const recordEdit = (outcome: ProjectionEdit) => outcomes.push(outcome);
          const before = session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
          if (trial === 4) editParagraphProperties(session, recordEdit);
          else edit(session, next, recordEdit);
          edits.push(...outcomes);
          const changed = session.storiesChangedSince(before).stories;
          const roots = new Set(changed.map(dirtyProjectionStory));
          const untouched = new Set([...canonicalStories(session, ts).keys()]
            .map(dirtyProjectionStory).filter((story) => !roots.has(story)));
          const label = `edit batch ${trial + 1} (${outcomes.map(({ kind, story, paraId, refusal }) =>
            `${kind}@${story}/${paraId} ${refusal ? `refused: ${refusal}` : 'applied'}`
          ).join(', ') || 'no eligible edit targets'})`;
          if (roots.size > 0) {
            const incrementalTs = yrsToDocument(session, ts, { storyIds: roots });
            const incrementalNative = nativeProjection(raw, buffer, roots);
            checkState(`${label}, selected roots`, check(session, incrementalTs, incrementalNative, source, roots));
            checkState(`${label}, untouched TypeScript roots`, compareSnapshots(session, ts, incrementalTs, untouched));
            checkState(`${label}, untouched native roots`, compareSnapshots(session, native.document, incrementalNative.document, untouched));

            ts = yrsToDocument(session, base);
            native = nativeProjection(raw, buffer);
            checkState(`${label}, full TypeScript projection`, compareSnapshots(session, ts, incrementalTs, undefined));
            checkState(`${label}, full native projection`, compareSnapshots(session, native.document, incrementalNative.document, undefined));
          } else {
            ts = yrsToDocument(session, base);
            native = nativeProjection(raw, buffer);
          }
          checkProjections(label, ts, native);
        }
        expect(edits.some((outcome) => !outcome.refusal), JSON.stringify(edits)).toBe(true);
        expect(failures, failures.join('\n\n')).toEqual([]);
      } finally {
        session.destroy();
      }
    }, 30_000);
  }

  afterAll(() => {
    const unused = projectionAllowlist.filter((entry) => !used.has(entry));
    const details = unused.map(({ family, kind, path, reason }) => `${family}/${kind} ${path}: ${reason}`);
    expect(unused, details.join('\n')).toEqual([]);
  });
});

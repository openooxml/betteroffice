import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import type { YrsSession } from './index';

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

export function normalizeProjection(value: unknown): unknown {
  if (value instanceof ArrayBuffer) return { base64: Buffer.from(value).toString('base64') };
  if (ArrayBuffer.isView(value)) {
    return { base64: Buffer.from(value.buffer, value.byteOffset, value.byteLength).toString('base64') };
  }
  if (value instanceof Date) return value.toISOString();
  if (value instanceof Map) {
    return Object.fromEntries([...value].map(([key, entry]) => [key, normalizeProjection(entry)]));
  }
  if (Array.isArray(value)) return value.map(normalizeProjection);
  const object = value !== null && typeof value === 'object'
    ? value as Record<string, unknown>
    : undefined;
  return object
    ? Object.fromEntries(Object.entries(object).filter(([, entry]) => entry !== undefined)
        .map(([key, entry]) => [key, normalizeProjection(entry)]))
    : value;
}

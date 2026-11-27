import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { recordRevisionIdCases } from './revisionIdsRecorder';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../../packages/docx/src/docx/rezip/parts';
import { writeDocumentWithRust } from '../../packages/docx/src/docx/rustSaveFacade';
import { markRevisionInfo } from '../../packages/docx/src/docx/sessionRevisionIds';
import type { ComplexField } from '../../packages/docx/src/types/content/link';
import type { Shape } from '../../packages/docx/src/types/content/shape';
import type { Document } from '../../packages/docx/src/types/document';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { preloadOpcWasm } from '../../packages/docx/src/wasm/opc';
import { createYrsSession, saveYrsDocx, type YrsLoc, type YrsSession } from '../../packages/docx/src/yrs';

const OUT = process.env.DOCX_REVISION_IDS_DIR;
const ROOT = resolve(import.meta.dir, '../..');
const SEED_ROOT = join(ROOT, 'packages/docx/src/yrs/__fixtures__/seed-parity');
const AUTHOR = { name: 'Parity Reviewer', date: '2026-01-01T00:00:00Z' };
const MARK_AUTHOR = { ...AUTHOR, name: 'Paragraph Mark Reviewer' };
const PROPERTY_AUTHOR = { ...AUTHOR, name: 'Paragraph Property Reviewer' };
const ROW_AUTHOR = { ...AUTHOR, name: 'Table Row Reviewer' };
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const DEFAULT_STYLES = `<w:styles xmlns:w="${W}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240" w:after="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
</w:styles>`;

const FIXTURE_ROOTS = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const FIXTURE_FILES = [
  'apps/demo/public/betteroffice-demo.docx',
  'crates/ooxml-opc/tests/fixtures/betteroffice-demo.docx',
  'crates/ooxml-redact/tests/fixtures/redaction-integrity.docx',
  'packages/docx-react/src/components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx',
];

const PART_TYPES: Record<string, string> = {
  'word/document.xml': 'wordprocessingml.document.main',
  'word/styles.xml': 'wordprocessingml.styles',
  'word/numbering.xml': 'wordprocessingml.numbering',
  'word/settings.xml': 'wordprocessingml.settings',
  'word/footnotes.xml': 'wordprocessingml.footnotes',
  'word/endnotes.xml': 'wordprocessingml.endnotes',
  'word/comments.xml': 'wordprocessingml.comments',
  'word/commentsExtended.xml': 'wordprocessingml.commentsExtended',
  'word/commentsIds.xml': 'wordprocessingml.commentsIds',
  'word/fontTable.xml': 'wordprocessingml.fontTable',
  'word/theme/theme1.xml': 'theme',
};

function relationships(entries: Array<[string, string, string]>): string {
  return '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('') +
    '</Relationships>';
}

function pack(parts: PartsMap): Uint8Array {
  if (!parts.has('_rels/.rels')) {
    parts.set('_rels/.rels', toBytes(relationships([['rIdDoc', 'officeDocument', 'word/document.xml']])));
  }
  if (!parts.has('[Content_Types].xml')) {
    const overrides = [...parts.keys()].flatMap((path) => {
      const kind = PART_TYPES[path] ??
        (/^word\/header\d+\.xml$/.test(path) ? 'wordprocessingml.header' : undefined) ??
        (/^word\/footer\d+\.xml$/.test(path) ? 'wordprocessingml.footer' : undefined);
      return kind ? [`<Override PartName="/${path}" ContentType="${OFFICE}.${kind}+xml"/>`] : [];
    });
    parts.set('[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' + overrides.join('') + '</Types>'
    ));
  }
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function files(dir: string): string[] {
  return readdirSync(dir, { recursive: true, encoding: 'utf8' }).sort();
}

function seedFixture(name: string): Uint8Array {
  const dir = join(SEED_ROOT, name);
  const parts: PartsMap = new Map();
  function readParts(prefix: string): void {
    const entries = readdirSync(join(dir, prefix), { withFileTypes: true });
    for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
      const part = prefix + entry.name;
      if (entry.isDirectory()) readParts(`${part}/`);
      else if (entry.isFile()) {
        const bytes = readFileSync(join(dir, part));
        if (part.endsWith('.base64')) {
          parts.set(part.slice(0, -'.base64'.length), Buffer.from(bytes.toString('utf8'), 'base64'));
        } else {
          parts.set(part, bytes);
        }
      }
    }
  }
  readParts('');
  if (!parts.has('word/styles.xml')) {
    parts.set('word/styles.xml', toBytes(DEFAULT_STYLES));
  }
  return pack(parts);
}

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const paragraph = (content: string) => `<w:p>${content}</w:p>`;

function synthetic(bookmarkId: number): Uint8Array {
  const plain = paragraph(run('Body text for revision edits.'));
  const retained = paragraph(
    `<w:ins w:id="7" w:author="Source">${run('Source revision')}</w:ins>` +
    `<w:bookmarkStart w:id="${bookmarkId}" w:name="SourceBookmark"/>${run('Bookmarked text')}<w:bookmarkEnd w:id="${bookmarkId}"/>` +
    '<w:commentRangeStart w:id="1"/>' + run('Commented') +
    '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>' +
    '<w:r><w:footnoteReference w:id="2"/></w:r>'
  );
  const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>' +
    paragraph(run('Cell text')) + '</w:tc></w:tr></w:tbl>';
  const formatting = paragraph(run('Paragraph formatting target.'));
  const properties = paragraph(run('Paragraph property target.'));
  const body = `${plain}${formatting}${properties}${retained}${table}<w:sectPr><w:headerReference w:type="default" r:id="h"/></w:sectPr>`;
  const xml = new Map([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`],
    ['word/_rels/document.xml.rels', relationships([
      ['h', 'header', 'header1.xml'],
      ['f', 'footnotes', 'footnotes.xml'],
      ['c', 'comments', 'comments.xml'],
    ])],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}">${paragraph(run('Header text'))}</w:hdr>`],
    ['word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="2">${paragraph('<w:r><w:footnoteRef/></w:r>' + run('Note text'))}</w:footnote></w:footnotes>`],
    ['word/comments.xml', `<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Source">${paragraph(run('Comment text'))}</w:comment></w:comments>`],
  ]);
  return pack(new Map([...xml].map(([path, value]) => [path, toBytes(value)])));
}

function editableText(session: YrsSession, story: string, characters: number, paraId?: string) {
  const spans = session.paragraphSpans(story);
  const indices = spans.flatMap((span, index) => paraId === undefined || span.paraId === paraId ? [index] : []);
  const units = session.storySegmentUnits(story, indices);
  for (const [index, unit] of units.entries()) {
    let offset = 0;
    for (const segment of unit) {
      if (segment.kind !== 'text') {
        offset += 1;
        continue;
      }
      const { attributes } = segment;
      if (!attributes.ins && !attributes.del && !attributes.rPrChange &&
        !attributes.fieldResult && !attributes.fieldCode && !attributes.hyperlink) {
        const head = Array.from(segment.text).slice(0, characters);
        if (head.length === characters) {
          return {
            at: { paraId: spans[indices[index]!]!.paraId, offset },
            lengths: head.map((character) => character.length),
            bold: attributes.bold === true,
          };
        }
      }
      offset += segment.text.length;
    }
  }
}

function replaceOne(session: YrsSession, story: string, target: NonNullable<ReturnType<typeof editableText>>) {
  const { at, lengths } = target;
  const end = { ...at, offset: at.offset + lengths[0]! };
  expect(session.replaceRange({ story, start: at, end }, 'Revised ', AUTHOR).revisionId).toBeTruthy();
}

async function scriptedSave(bytes: Uint8Array, edit: (session: YrsSession) => void) {
  const session = await createYrsSession({ clientId: 3 });
  try {
    session.openDocx(bytes, true);
    edit(session);
    await saveYrsDocx(session, { updateModifiedDate: false });
  } finally {
    session.destroy();
  }
}

async function saveStructuredFields() {
  const revisedRun = (id: number, text: string) => ({
    type: 'run' as const,
    propertyChanges: [{
      type: 'runPropertyChange' as const,
      info: markRevisionInfo({ id, author: AUTHOR.name, date: AUTHOR.date }, true),
      previousFormatting: { bold: true },
    }],
    content: [{ type: 'text' as const, text }],
  });
  const field = (placeholder: number): ComplexField => ({
    type: 'complexField',
    instruction: 'QUOTE',
    fieldType: 'QUOTE',
    fieldCode: [],
    fieldResult: [],
    structuredResult: {
      inline: [revisedRun(placeholder, 'Inline result')],
      blocks: [{ type: 'paragraph', content: [revisedRun(placeholder - 1, 'Block result')] }],
    },
  });
  const textBox: Shape & { textBox: true } = {
    type: 'shape',
    shapeType: 'rect',
    textBox: true,
    size: { width: 914400, height: 457200 },
    textBody: { content: [{ type: 'paragraph', content: [field(Number.MAX_SAFE_INTEGER - 2)] }] },
  };
  const document: Document = {
    package: { document: { content: [
      { type: 'paragraph', content: [field(Number.MAX_SAFE_INTEGER)] },
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'shape', shape: textBox }] }] },
    ] } },
  };
  await writeSyntheticDocument(document);
}

async function writeSyntheticDocument(document: Document) {
  const source = pack(new Map([['word/document.xml', toBytes(
    `<w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`
  )]]));
  await writeDocumentWithRust(document, Uint8Array.from(source).buffer, { updateModifiedDate: false });
}

async function saveArrayBufferViews() {
  const placeholder = 2 ** 40;
  const payload = Buffer.from([1]);
  const withXml = Object.assign(Buffer.from([2]), {
    toJSON: () => ({
      type: 'Buffer', data: [2], xml: "w:id='90'", info: { id: placeholder, author: AUTHOR.name },
    }),
  });
  const shape: Shape = {
    type: 'shape',
    shapeType: 'rect',
    size: { width: 914400, height: 457200 },
    drawingModel: {
      frame: { 'payload/~': [payload, withXml], record: { type: 'Buffer', data: [1] } },
      otherSha256: '',
    },
  };
  await writeSyntheticDocument({ package: { document: { content: [
    { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'shape', shape }] }] },
    { type: 'paragraph', content: [{
      type: 'insertion',
      info: markRevisionInfo({ id: placeholder, author: AUTHOR.name, date: AUTHOR.date }, true),
      content: [],
    }] },
  ] } } });
}

function bodyEdits(session: YrsSession, mixed: boolean) {
  const story = 'body';
  let target = editableText(session, story, 3);
  if (!target) {
    const last = session.paragraphSpans(story).at(-1);
    if (!last) throw new Error('no body paragraph for the revision script');
    session.insertText({ story, paraId: last.paraId, offset: last.length }, 'Parity text');
    target = editableText(session, story, 3);
    if (!target) throw new Error('no editable body text for the revision script');
  }
  const { at, lengths } = target;
  const split = session.splitParagraph({ story, ...at, offset: at.offset + lengths[0]! + lengths[1]! }, MARK_AUTHOR);
  expect(split.revisionId).toBeTruthy();
  expect(split.firstParaId).toBe(at.paraId);

  if (mixed) {
    formatOne(session, story, target, 1);
    replaceOne(session, story, target);
    paragraphEdit(session, { story, ...at });
    expect(session.insertText({ story, paraId: split.secondParaId, offset: 0 }, 'Following ', AUTHOR).revisionId).toBeTruthy();
    return;
  }

  const paragraphs = session.paragraphs(story);
  const index = paragraphs.findIndex(({ paraId }) => paraId === split.secondParaId);
  const [replacement, formatting, properties] = paragraphs.slice(index, index + 3);
  if (!replacement || !formatting || !properties) throw new Error('missing separate revision paragraphs');
  const replace = editableText(session, story, 1, replacement.paraId);
  const format = editableText(session, story, 1, formatting.paraId);
  if (!replace || !format) throw new Error('missing separate revision text');
  replaceOne(session, story, replace);
  formatOne(session, story, format, 0);
  paragraphEdit(session, { story, paraId: properties.paraId, offset: 0 });
}

function formatOne(session: YrsSession, story: string, target: NonNullable<ReturnType<typeof editableText>>, character: 0 | 1) {
  const { at, lengths, bold } = target;
  const start = { ...at, offset: at.offset + (character === 1 ? lengths[0]! : 0) };
  const end = { ...start, offset: start.offset + lengths[character]! };
  session.formatRange({ story, start, end }, { bold: !bold }, AUTHOR);
  expect(session.listRevisions().some(({ kind, author }) => kind === 'rPrChange' && author === AUTHOR.name)).toBe(true);
}

function nestedEdits(session: YrsSession, requireNested: boolean) {
  const stories = session.listStories().stories;
  for (const kinds of [['tableCell'], ['header'], ['footnote', 'endnote']]) {
    let changed = false;
    for (const { kind, story } of stories) {
      if (!kinds.includes(kind)) continue;
      const target = editableText(session, story, 1);
      if (!target) continue;
      replaceOne(session, story, target);
      changed = true;
      break;
    }
    if (requireNested && !changed) throw new Error(`no editable ${kinds.join('/')} story`);
  }
}

function paragraphEdit(session: YrsSession, at: YrsLoc) {
  const paragraph = session.paragraphs(at.story).find(({ paraId }) => paraId === at.paraId);
  const alignment = paragraph?.properties.alignment === 'right' ? 'center' : 'right';
  session.setParagraphAttrs({ story: at.story, start: at, end: at }, { alignment }, PROPERTY_AUTHOR);
  expect(session.listRevisions().some(({ kind, author }) => kind === 'pPrChange' && author === PROPERTY_AUTHOR.name)).toBe(true);
}

function tableEdits(session: YrsSession) {
  const story = 'body';
  const inserted = session.insertRow({ story, tableIndex: 0, row: 0, column: 0 }, 'below', ROW_AUTHOR);
  expect(inserted.revisionIds).toHaveLength(1);
  const cellStory = inserted.createdStoryIds[0]!;
  const [cell] = session.paragraphSpans(cellStory);
  expect(session.insertText({ story: cellStory, paraId: cell!.paraId, offset: 0 }, 'Inserted cell text', AUTHOR).revisionId).toBeTruthy();

  let index = 0;
  for (const segment of session.storySegments(story)) {
    if (segment.kind === 'embed' && segment.embedKind === 'table') {
      const rows = segment.payload.rows as Array<{
        cells: Array<{ story: string; tcPr?: Record<string, unknown> }>;
      }>;
      const target = rows.flatMap(({ cells }) => cells).find(({ story }) => story === cellStory);
      if (!target) throw new Error('no inserted table cell');
      target.tcPr = {
        ...target.tcPr,
        cellMarker: { kind: 'ins', info: { id: 'parity-cell-insertion', author: AUTHOR.name, date: AUTHOR.date } },
      };
      session.applyRawOps(story, [{ op: 'setEmbedAttr', index, key: 'rows', value: rows }]);
      return;
    }
    index += segment.kind === 'text' ? segment.text.length : 1;
  }
  throw new Error('no scripted table');
}

describe.skipIf(!OUT)('session revision IDs for the Rust save request', () => {
  if (!OUT) return;

  beforeAll(async () => {
    const wasm = join(ROOT, 'packages/docx/src/wasm/generated');
    await preloadEditWasm(new Uint8Array(readFileSync(join(wasm, 'edit/docx_edit_bg.wasm'))));
    await preloadOpcWasm(new Uint8Array(readFileSync(join(wasm, 'opc/ooxml_opc_bg.wasm'))));
  }, 30_000);

  function scriptedDocument(name: string, load: () => Uint8Array, synthetic = false, mixed = true) {
    test(name, async () => {
      expect(await recordRevisionIdCases(name, () => scriptedSave(load(), (session) => {
        bodyEdits(session, mixed);
        nestedEdits(session, synthetic);
        if (synthetic) tableEdits(session);
      }))).toBeGreaterThan(0);
    }, 120_000);
  }

  const paths = new Set([
    ...FIXTURE_ROOTS.flatMap((dir) =>
      files(join(ROOT, dir)).filter((path) => path.endsWith('.docx')).map((path) => join(dir, path))
    ),
    ...FIXTURE_FILES,
  ]);
  for (const path of [...paths].sort()) {
    scriptedDocument(path, () => new Uint8Array(readFileSync(join(ROOT, path))));
  }
  for (const name of readdirSync(SEED_ROOT, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort()) {
    scriptedDocument(`seed-parity-${name}`, () => seedFixture(name));
  }
  scriptedDocument('synthetic-source-annotations', () => synthetic(120), true);
  scriptedDocument('synthetic-wrap', () => synthetic(2 ** 31 - 1), true);
  scriptedDocument('synthetic-separate-paragraphs', () => synthetic(120), true, false);
  test('synthetic-structured-field-results', async () => {
    expect(await recordRevisionIdCases('synthetic-structured-field-results', saveStructuredFields)).toBe(1);
  }, 30_000);
  test('synthetic-array-buffer-views', async () => {
    expect(await recordRevisionIdCases('synthetic-array-buffer-views', saveArrayBufferViews)).toBe(1);
  }, 30_000);
});

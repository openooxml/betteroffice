import { afterEach, beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { DirtyProjectionStories, mergeDocxHostMetadata, saveEditorDocument } from './editorSave';
import { createYrsSession, decodeDocxHostJson, saveYrsDocx, type YrsSession } from './index';
import { closeCommandEngines, openCommandEngine } from './residentCommandTestSupport';
import * as residentEngine from './residentEngineSession';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NAMESPACES = [
  `xmlns:w="${W}"`,
  `xmlns:r="${R}"`,
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
].join(' ');
const AUTHOR = 'Grace';
const DATE = '2025-03-04T05:06:07Z';
const HEAD = 'Head ';

const PATHS = ['Rust loader', 'TypeScript loader', 'worker engine'] as const;
type Path = (typeof PATHS)[number];

const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const dt = (text: string) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`;
const textRun = (deleted: boolean, text: string) => (deleted ? dt(text) : t(text));
const tracked = (tag: string, id: string, content: string) =>
  `<w:${tag} w:id="${id}" w:author="${AUTHOR}" w:date="${DATE}">${content}</w:${tag}>`;
const link = (content: string, attributes = 'r:id="rIdLink"') =>
  `<w:hyperlink ${attributes} w:history="1">${content}</w:hyperlink>`;
const FIRST = '<w:p w14:paraId="0C000001" w14:textId="0C000001">';
const firstParagraph = (inner: string) => `${FIRST}${t(HEAD)}${inner}${t(' tail')}</w:p>`;
const OTHER = `<w:p w14:paraId="0C000002" w14:textId="0C000002">${t('Other')}</w:p>`;

const preserved = (text: string) => (/^ | $| {2}/.test(text) ? ' xml:space="preserve"' : '');
/** A plain run as the writer writes it. */
const wt = (text: string) => `<w:r><w:t${preserved(text)}>${text}</w:t></w:r>`;

type Notes = 'footnotes' | 'endnotes';

interface Parts {
  notes?: Notes;
  /** A comments part holding comment 7. */
  comments?: boolean;
  relationships?: string;
}

function docx(body: string, { notes, comments = false, relationships = '' }: Parts = {}): Uint8Array {
  const parts: Array<[string, string]> = [
    [
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.document.main+xml"/>` +
        (notes ? `<Override PartName="/word/${notes}.xml" ContentType="${OFFICE_DOC}.${notes}+xml"/>` : '') +
        (comments ? `<Override PartName="/word/comments.xml" ContentType="${OFFICE_DOC}.comments+xml"/>` : '') +
        `<Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.styles+xml"/></Types>`,
    ],
    [
      '_rels/.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ],
    [
      'word/_rels/document.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` +
        (notes ? `<Relationship Id="rIdNotes" Type="${R}/${notes}" Target="${notes}.xml"/>` : '') +
        (comments ? `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/>` : '') +
        `<Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/>${relationships}</Relationships>`,
    ],
    [
      'word/styles.xml',
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    ],
    ['word/document.xml', `<w:document ${NAMESPACES}><w:body>${body}<w:sectPr/></w:body></w:document>`],
  ];
  if (comments) {
    parts.push([
      'word/comments.xml',
      `<w:comments xmlns:w="${W}">` +
        `<w:comment w:id="7" w:author="${AUTHOR}" w:date="${DATE}" w:initials="G"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>`,
    ]);
  }
  if (notes) {
    const note = notes === 'footnotes' ? 'footnote' : 'endnote';
    parts.push([
      `word/${notes}.xml`,
      `<w:${notes} xmlns:w="${W}">` +
        `<w:${note} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${note}>` +
        `<w:${note} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${note}>` +
        `<w:${note} w:id="1"><w:p><w:r><w:${note}Ref/></w:r><w:r><w:t xml:space="preserve"> Note</w:t></w:r></w:p></w:${note}></w:${notes}>`,
    ]);
  }
  return new Uint8Array(rezipPartsToArrayBuffer(new Map(parts.map(([name, xml]) => [name, toBytes(xml)]))));
}

const documentXml = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
const body = (xml: string) => xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'));
/** The saved first body paragraph, start tag included. */
const savedFirst = (xml: string) => body(xml).match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/)?.[0] ?? '';
/** The saved first body paragraph without its start tag and end tag. */
const savedFirstInner = (xml: string) =>
  savedFirst(xml)
    .replace(/^<w:p\b[^>]*>/, '')
    .replace(/<\/w:p>$/, '');

interface Opened {
  session: Pick<YrsSession, 'paragraphs' | 'insertText'>;
  /** The path writes every paragraph again on each save: it has no source package. */
  rewrites: boolean;
  revisions(): ReadonlyArray<{ revisionId: string; kind: string }>;
  saves(): Promise<Array<[string, Uint8Array]>>;
}

const sessions: YrsSession[] = [];
let clientId = 79600;

beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  );
  await preloadOpcWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')))
  );
}, 30_000);
afterEach(() => {
  sessions.splice(0).forEach((session) => session.destroy());
  closeCommandEngines();
});

async function yrs(): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: (clientId += 1) });
  sessions.push(session);
  return session;
}

/** The comments the host holds for the opened document, which the editor hands to its save. */
const hostComments = (document: Document): Comment[] => document.package.document.comments ?? [];

/** The editor's main-thread save: the body projected onto `base`, merged with the opening host when there is one. */
async function editorSave(session: YrsSession, base: Document, comments: Comment[], host?: Document): Promise<Uint8Array> {
  const dirty = new DirtyProjectionStories();
  dirty.add('body');
  const projected = yrsToDocument(session, host ? mergeDocxHostMetadata(base, host) : base, dirty.projectionOptions());
  return new Uint8Array(await saveEditorDocument(session, projected, comments));
}

async function open(path: Path, bytes: Uint8Array): Promise<Opened> {
  if (path === 'Rust loader') {
    const session = await yrs();
    const host = session.openDocx(bytes, true).document;
    return {
      session,
      rewrites: false,
      revisions: () => session.listRevisions(),
      saves: async () => [
        ['saveYrsDocx', (await saveYrsDocx(session, { updateModifiedDate: false })).bytes],
        ['editor', await editorSave(session, session.materializeDocx()!, hostComments(host), host)],
      ],
    };
  }
  if (path === 'TypeScript loader') {
    const session = await yrs();
    const parsed = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
    documentToYrs(session, parsed);
    const base = { ...parsed, originalBuffer: parsed.originalBuffer ?? bytes.slice().buffer };
    return {
      session,
      rewrites: true,
      revisions: () => session.listRevisions(),
      saves: async () => [['editor', await editorSave(session, base, hostComments(parsed))]],
    };
  }
  let hostJson = '';
  const create = residentEngine.createResidentEngineSession;
  const capture = spyOn(residentEngine, 'createResidentEngineSession').mockImplementation(async (...args) => {
    const session = await create(...args);
    const openDocx = session.openDocx;
    session.openDocx = (...input) => (hostJson = openDocx(...input));
    return session;
  });
  const { engine } = await openCommandEngine([], { bytes, clientId: (clientId += 1) }).finally(() =>
    capture.mockRestore()
  );
  return {
    session: engine,
    rewrites: false,
    revisions: () => engine.proposalEngine.listRevisions(),
    saves: async () => {
      const comments = hostComments(decodeDocxHostJson(hostJson, bytes).document);
      return [
        ['worker', new Uint8Array(await engine.save(bytes.slice(), hostJson, undefined, comments, { full: false }))],
      ];
    },
  };
}

function typedAt({ session }: Opened, paragraph: number, offset: number): void {
  const { paraId } = session.paragraphs('body')[paragraph]!;
  session.insertText({ story: 'body', paraId, offset }, 'X');
}

const texts = ({ session }: Opened) => session.paragraphs('body').map(({ text }) => text);

interface Change {
  tag: 'ins' | 'del' | 'moveFrom' | 'moveTo';
  id: string;
  deleted: boolean;
}

interface Side {
  name: string;
  change: Change;
}

const INSERTION: Side = { name: 'an insertion', change: { tag: 'ins', id: '11', deleted: false } };
const DELETION: Side = { name: 'a deletion', change: { tag: 'del', id: '12', deleted: true } };
const MOVED_FROM: Side = { name: 'a moved-from change', change: { tag: 'moveFrom', id: '13', deleted: true } };
const MOVED_TO: Side = { name: 'a moved-to change', change: { tag: 'moveTo', id: '14', deleted: false } };

interface Kind {
  name: string;
  notes?: Notes;
  /** The document holds comment 7, whose reference run follows the hyperlink. */
  comments?: boolean;
  sides?: Side[];
  content(deleted: boolean): string;
}

const COMMENT_REFERENCE = '<w:r><w:commentReference w:id="7"/></w:r>';
const afterLink = (kind: Kind) => (kind.comments ? COMMENT_REFERENCE : '');

const DRAWING =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="9525" cy="9525"/>' +
  '<wp:docPr id="1" name="Shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"/>' +
  '</a:graphic></wp:inline></w:drawing></w:r>';

const TEXT_KINDS: Kind[] = [
  {
    name: 'text',
    sides: [INSERTION, DELETION, MOVED_FROM, MOVED_TO],
    content: (deleted) => textRun(deleted, 'added'),
  },
  { name: 'a tab', content: () => '<w:r><w:tab/></w:r>' },
  { name: 'a line break', content: () => '<w:r><w:br/></w:r>' },
];

const KINDS: Kind[] = [
  { name: 'a footnote reference', notes: 'footnotes', content: () => '<w:r><w:footnoteReference w:id="1"/></w:r>' },
  { name: 'an endnote reference', notes: 'endnotes', content: () => '<w:r><w:endnoteReference w:id="1"/></w:r>' },
  {
    name: 'an inline content control',
    content: (deleted) =>
      `<w:sdt><w:sdtPr><w:alias w:val="Box"/></w:sdtPr><w:sdtContent>${textRun(deleted, 'boxed')}</w:sdtContent></w:sdt>`,
  },
  { name: 'a page break', content: () => '<w:r><w:br w:type="page"/></w:r>' },
  { name: 'a column break', content: () => '<w:r><w:br w:type="column"/></w:r>' },
  { name: 'a simple field', content: (deleted) => `<w:fldSimple w:instr=" PAGE ">${textRun(deleted, '1')}</w:fldSimple>` },
  {
    name: 'a complex field',
    content: (deleted) => {
      const instruction = deleted ? 'delInstrText' : 'instrText';
      return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        `<w:r><w:${instruction} xml:space="preserve"> PAGE </w:${instruction}></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${textRun(deleted, '2')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
      );
    },
  },
  { name: 'a drawing', content: () => DRAWING },
  { name: 'a symbol', content: () => '<w:r><w:sym w:font="Wingdings" w:char="F04A"/></w:r>' },
  { name: 'an equation', content: () => '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>' },
  {
    name: 'a comment range mark',
    comments: true,
    content: (deleted) => `<w:commentRangeStart w:id="7"/>${textRun(deleted, 'marked')}<w:commentRangeEnd w:id="7"/>`,
  },
  {
    name: 'a bookmark mark',
    content: (deleted) => `<w:bookmarkStart w:id="5" w:name="mark"/>${textRun(deleted, 'marked')}<w:bookmarkEnd w:id="5"/>`,
  },
  {
    name: 'a nested tracked change',
    content: (deleted) =>
      deleted ? `${dt('gone')}${tracked('ins', '21', t('added'))}` : `${t('kept')}${tracked('del', '22', dt('dropped'))}`,
  },
];

const MIXED: Kind = {
  name: 'text and a footnote reference',
  notes: 'footnotes',
  content: (deleted) =>
    `${textRun(deleted, 'lost')}<w:r><w:footnoteReference w:id="1"/></w:r>${textRun(deleted, 'more')}`,
};

/** The paragraph written without the tracked change: the link keeps its remaining text as two plain runs. */
const rewritten = (head: string, kind: Kind) =>
  `${wt(head)}${link(`${wt('see ')}${wt(' here')}`)}${afterLink(kind)}${wt(' tail')}`;

type Cell = [string, Path, Kind, Change];

const cells = (kinds: Kind[]): Cell[] =>
  kinds.flatMap((kind) =>
    (kind.sides ?? [INSERTION, DELETION]).flatMap((side) =>
      PATHS.map((path) => [`${kind.name} in ${side.name}`, path, kind, side.change] as Cell)
    )
  );

const KEPT = 'keeps its paragraph as opened';
const DROPPED = 'is written without the tracked change';
const keepsSource = (path: Path, kind: Kind) => path !== 'TypeScript loader' && !kind.comments;

const elsewhereCells = (kinds: Kind[]) =>
  cells(kinds).map(
    ([name, path, kind, change]) =>
      [name, path, keepsSource(path, kind) ? KEPT : DROPPED, kind, change] as [string, Path, string, Kind, Change]
  );

const source = (kind: Kind, change: Change) =>
  firstParagraph(
    `${link(`${t('see ')}${tracked(change.tag, change.id, kind.content(change.deleted))}${t(' here')}`)}${afterLink(kind)}`
  );

const packageOf = (kind: Kind, paragraph: string) =>
  docx(`${paragraph}${OTHER}`, { notes: kind.notes, comments: kind.comments });

function describeInsideLink(title: string, kinds: Kind[]): void {
  describe(title, () => {
    it.each(cells(kinds))('%s inside a hyperlink, opened by the %s, is not loaded', async (_name, path, kind, change) => {
      const opened = await open(path, packageOf(kind, source(kind, change)));
      expect(opened.revisions()).toEqual([]);
      expect(texts(opened)).toEqual([`${HEAD}see  here tail`, 'Other']);
    });

    it.each(elsewhereCells(kinds))(
      '%s inside a hyperlink, opened by the %s, %s after an edit in another paragraph',
      async (_name, path, _outcome, kind, change) => {
        const paragraph = source(kind, change);
        const opened = await open(path, packageOf(kind, paragraph));
        typedAt(opened, 1, 0);
        for (const [save, saved] of await opened.saves()) {
          const xml = documentXml(saved);
          if (opened.rewrites) expect([save, savedFirstInner(xml)]).toEqual([save, rewritten(HEAD, kind)]);
          else
            expect([save, savedFirst(xml)]).toEqual([
              save,
              keepsSource(path, kind) ? paragraph : `${FIRST}${rewritten(HEAD, kind)}</w:p>`,
            ]);
        }
      }
    );

    it.each(cells(kinds))(
      '%s inside a hyperlink, opened by the %s, is written without the tracked change after an edit in that paragraph',
      async (_name, path, kind, change) => {
        const opened = await open(path, packageOf(kind, source(kind, change)));
        typedAt(opened, 0, 2);
        for (const [save, saved] of await opened.saves()) {
          expect([save, savedFirstInner(documentXml(saved))]).toEqual([save, rewritten('HeXad ', kind)]);
        }
      }
    );
  });
}

describeInsideLink('a tracked change inside a hyperlink that holds text', TEXT_KINDS);
describeInsideLink('a tracked change inside a hyperlink that holds more than text', KINDS);
describeInsideLink('a tracked change inside a hyperlink that holds text and a footnote reference', [MIXED]);

describe('two adjacent untracked hyperlinks to the same address', () => {
  const TOOLTIPS = `<Relationship Id="rIdLink2" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/>`;
  const inner = `${link(t('first'), 'r:id="rIdLink" w:tooltip="First"')}${link(t('second'), 'r:id="rIdLink2" w:tooltip="Second"')}`;
  const HYPERLINK = /<w:hyperlink\b((?:[^>/]|\/(?!>))*)>([\s\S]*?)<\/w:hyperlink>/g;
  const attribute = (attributes: string, name: string) => attributes.match(new RegExp(`${name}="([^"]*)"`))?.[1];
  const linked = (xml: string) =>
    [...body(xml).matchAll(HYPERLINK)].map(([, attributes, content]) => [
      attribute(attributes!, 'r:id'),
      attribute(attributes!, 'w:tooltip'),
      [...content!.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(([, text]) => text).join(''),
    ]);

  const TWO = [
    ['rIdLink', 'First', 'first'],
    ['rIdLink2', 'Second', 'second'],
  ];
  const ONE = [['rIdLink', 'First', 'firstsecond']];
  const WRITTEN: Array<[Path, string, string[][]]> = [
    ['Rust loader', 'two links', TWO],
    ['TypeScript loader', 'one link under the first link', ONE],
    ['worker engine', 'two links', TWO],
  ];

  it.each(WRITTEN)('opened by the %s, are written as %s after an edit in their paragraph', async (path, _shape, written) => {
    const opened = await open(path, docx(`${firstParagraph(inner)}${OTHER}`, { relationships: TOOLTIPS }));
    typedAt(opened, 0, 2);
    for (const [save, saved] of await opened.saves()) {
      expect([save, linked(documentXml(saved))]).toEqual([save, written]);
    }
  });
});

describe('a deleted page break that ends a hyperlink and its paragraph', () => {
  const DELETED_BREAK = tracked('del', '12', '<w:r><w:br w:type="page"/></w:r>');
  const paragraph = `${FIRST}${t(HEAD)}${link(`${t('see ')}${DELETED_BREAK}`)}</w:p>`;
  const BREAKS = /<w:br\b|<w:pageBreakBefore\b/g;
  /** What the save writes after the first paragraph: the count of paragraphs and of breaks of any kind. */
  const afterFirst = (xml: string) => {
    const rest = body(xml).slice(body(xml).indexOf(savedFirst(xml)) + savedFirst(xml).length);
    return [rest.match(/<w:p\b/g)?.length ?? 0, rest.match(BREAKS)?.length ?? 0];
  };
  const byPath = PATHS.map((path) => [path] as [Path]);
  const elsewhere = PATHS.map(
    (path) => [path, path === 'TypeScript loader' ? 'is written without the break' : 'keeps its paragraph as opened'] as [Path, string]
  );

  it.each(byPath)('opened by the %s, is not loaded', async (path) => {
    const opened = await open(path, docx(`${paragraph}${OTHER}`));
    expect(opened.revisions()).toEqual([]);
    expect(texts(opened)).toEqual([`${HEAD}see `, 'Other']);
  });

  it.each(elsewhere)('opened by the %s, %s after an edit in another paragraph', async (path) => {
    const opened = await open(path, docx(`${paragraph}${OTHER}`));
    typedAt(opened, 1, 0);
    for (const [save, saved] of await opened.saves()) {
      const xml = documentXml(saved);
      const first = opened.rewrites ? savedFirstInner(xml) : savedFirst(xml);
      expect([save, first, afterFirst(xml)]).toEqual([
        save,
        opened.rewrites ? `${wt(HEAD)}${link(wt('see '))}` : paragraph,
        [1, 0],
      ]);
    }
  });

  it.each(byPath)('opened by the %s, is written without the break after an edit in that paragraph', async (path) => {
    const opened = await open(path, docx(`${paragraph}${OTHER}`));
    typedAt(opened, 0, 2);
    for (const [save, saved] of await opened.saves()) {
      const xml = documentXml(saved);
      expect([save, savedFirstInner(xml), afterFirst(xml)]).toEqual([save, `${wt('HeXad ')}${link(wt('see '))}`, [1, 0]]);
    }
  });
});

import { beforeAll, describe, expect, it, spyOn } from 'bun:test';
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
import { createResidentEngineSession } from './residentEngineSession';
import type { ResidentSaveRecord } from './residentSave';
import { yrsToDocument } from './yrsToDocument';

const TIMEOUT = 60_000;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NAMESPACES = [
  `xmlns:w="${W}"`,
  `xmlns:r="${R}"`,
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(' ');

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const start = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
const end = (id: number) => `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
let paragraphCount = 0;
const para = (inner: string) => {
  paragraphCount += 1;
  const id = `0F${String(paragraphCount).padStart(6, '0')}`;
  return `<w:p w14:paraId="${id}" w14:textId="${id}">${inner}</w:p>`;
};
const table = (...cells: string[]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid>' +
  cells.map(() => '<w:gridCol w:w="2500"/>').join('') +
  '</w:tblGrid><w:tr>' +
  cells.map((cell) => `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>${cell}</w:tc>`).join('') +
  '</w:tr></w:tbl>';
const textBox = (inner: string) =>
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="20" name="Text Box"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvPr id="21" name="Shape 21"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>' +
  `</a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx>` +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

function docx(body: string, ids: number[], header?: string): Uint8Array {
  const comment = (id: number) =>
    `<w:comment w:id="${id}" w:author="Grace" w:date="2025-03-04T05:06:07Z" w:initials="G">` +
    `<w:p><w:r><w:t>Note ${id}</w:t></w:r></w:p></w:comment>`;
  const parts: Array<[string, string]> = [
    [
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/>` +
        `<Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/>` +
        (header === undefined ? '' : `<Override PartName="/word/header1.xml" ContentType="${OFFICE}.header+xml"/>`) +
        `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`,
    ],
    [
      '_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ],
    [
      'word/_rels/document.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` +
        (header === undefined ? '' : `<Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/>`) +
        `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/></Relationships>`,
    ],
    [
      'word/styles.xml',
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
        '<w:name w:val="Normal"/></w:style></w:styles>',
    ],
    ['word/comments.xml', `<w:comments xmlns:w="${W}">${ids.map(comment).join('')}</w:comments>`],
    [
      'word/document.xml',
      `<w:document ${NAMESPACES}><w:body>${body}` +
        (header === undefined ? '<w:sectPr/>' : '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>') +
        '</w:body></w:document>',
    ],
    ...(header === undefined ? [] : [['word/header1.xml', `<w:hdr ${NAMESPACES}>${header}</w:hdr>`] as [string, string]]),
  ];
  return new Uint8Array(rezipPartsToArrayBuffer(new Map(parts.map(([name, xml]) => [name, toBytes(xml)]))));
}

const OPENING = `${run('Alpha ')}${start(1)}${start(2)}${run('one')}${end(2)}${run('.')}`;
const ENDING = `${run('Gamma')}${end(1)}${run(' three.')}`;
const CELL_A = 'body:t0:r0c0';
const CELL_B = 'body:t0:r0c1';

type ActionName =
  | 'no edit'
  | 'typing where the range starts'
  | 'typing in the next paragraph'
  | 'typing where the range ends'
  | 'bold where the range starts'
  | 'deleting another comment'
  | 'deleting the comment';

interface Typing {
  story: string;
  paragraph: number;
  /** An offset in the paragraph, or its end. */
  offset: number | 'end';
  /** What the commented text reads once typed. */
  text?: string;
}

interface Built {
  name: string;
  bytes: Uint8Array;
  /** The comments whose ranges the cells follow. */
  ids: number[];
  other: number;
  last: string;
  bold: { story: string; to: number };
  typing: Partial<Record<ActionName, Typing>>;
  /** The stories each comment is shown in after opening. */
  anchors: Record<number, string[]>;
  /** The comments the session lists. */
  listed: string[];
}

const BUILT: Built[] = [
  {
    name: 'a comment that ends in a table cell',
    bytes: docx(
      para(OPENING) + para(run('Beta two.')) + table(para(ENDING)) + para(run('Delta four.')),
      [1, 2]
    ),
    ids: [1],
    other: 2,
    last: 'Delta four.',
    bold: { story: 'body', to: 5 },
    typing: {
      'typing where the range starts': { story: 'body', paragraph: 0, offset: 7, text: 'oXne.Beta two.Gamma' },
      'typing in the next paragraph': { story: 'body', paragraph: 1, offset: 5, text: 'one.Beta Xtwo.Gamma' },
      'typing where the range ends': { story: CELL_A, paragraph: 0, offset: 2, text: 'one.Beta two.GaXmma' },
    },
    anchors: { 1: ['body'] },
    listed: ['1', '2'],
  },
  {
    name: 'a comment that ends in a later body paragraph',
    bytes: docx(para(OPENING) + para(run('Beta two.')) + para(ENDING) + para(run('Delta four.')), [1, 2]),
    ids: [1],
    other: 2,
    last: 'Delta four.',
    bold: { story: 'body', to: 5 },
    typing: {
      'typing where the range starts': { story: 'body', paragraph: 0, offset: 7, text: 'oXne.Beta two.Gamma' },
      'typing in the next paragraph': { story: 'body', paragraph: 1, offset: 5, text: 'one.Beta Xtwo.Gamma' },
      'typing where the range ends': { story: 'body', paragraph: 2, offset: 2, text: 'one.Beta two.GaXmma' },
    },
    anchors: { 1: ['body'] },
    listed: ['1', '2'],
  },
  {
    name: 'a comment that ends in a text box',
    bytes: docx(
      para(OPENING) + para(run('Beta two.')) + para(`${run('Box ')}${textBox(para(ENDING))}`) + para(run('Delta four.')),
      [1, 2]
    ),
    ids: [1],
    other: 2,
    last: 'Delta four.',
    bold: { story: 'body', to: 5 },
    typing: {
      'typing where the range starts': { story: 'body', paragraph: 0, offset: 7, text: 'oXne.Beta two.Box Gamma' },
      'typing in the next paragraph': { story: 'body', paragraph: 1, offset: 5, text: 'one.Beta Xtwo.Box Gamma' },
    },
    anchors: { 1: ['body'] },
    listed: ['1', '2'],
  },
  {
    name: 'a comment from one cell to the start of the next',
    bytes: docx(
      para(`${start(2)}${run('Before')}${end(2)}${run(' row.')}`) +
        table(para(`${start(1)}${run('Left cell text.')}`), para(`${end(1)}${run('Right cell text.')}`)) +
        para(run('After row.')),
      [1, 2]
    ),
    ids: [1],
    other: 2,
    last: 'After row.',
    bold: { story: CELL_A, to: 4 },
    typing: {
      'typing where the range starts': { story: CELL_A, paragraph: 0, offset: 2, text: 'LeXft cell text.' },
      'typing where the range ends': { story: CELL_B, paragraph: 0, offset: 'end' },
    },
    anchors: { 1: [CELL_A] },
    listed: ['1', '2'],
  },
  {
    name: 'three comments from one cell to the start of the next',
    bytes: docx(
      para(`${start(2)}${run('Before')}${end(2)}${run(' row.')}`) +
        table(
          para(`${start(1)}${start(4)}${start(5)}${run('Left cell text.')}`),
          para(`${end(1)}${end(4)}${end(5)}${run('Right cell text.')}`)
        ) +
        para(run('After row.')),
      [1, 2, 4, 5]
    ),
    ids: [1, 4, 5],
    other: 2,
    last: 'After row.',
    bold: { story: CELL_A, to: 4 },
    typing: {
      'typing where the range starts': { story: CELL_A, paragraph: 0, offset: 2, text: 'LeXft cell text.' },
      'typing where the range ends': { story: CELL_B, paragraph: 0, offset: 'end' },
    },
    anchors: { 1: [CELL_A], 4: [CELL_A], 5: [CELL_A] },
    listed: ['1', '2', '4', '5'],
  },
];

const ROUTES = ['path that seeds in Rust', 'worker engine', 'path that seeds in TypeScript'] as const;
type Route = (typeof ROUTES)[number];
const KEEPS_BYTES = (route: Route) => route !== 'path that seeds in TypeScript';

type Editable = Pick<
  YrsSession,
  | 'paragraphs'
  | 'insertText'
  | 'toggleMark'
  | 'applyRawOps'
  | 'resolveComment'
  | 'listComments'
  | 'deleteRange'
  | 'deleteColumn'
  | 'mergeCells'
  | 'beginUndoCapture'
  | 'addUndoBoundary'
  | 'undo'
>;

interface Opened {
  session: Editable;
  /** The length of each paragraph of the story, in the session's offsets. */
  lengths: YrsSession['paragraphSpans'];
  revisions(): number;
  /** Saves as the route's editor does; `stories` are marked changed besides the body. */
  save(comments: (list: Comment[]) => Comment[], stories?: readonly string[]): Promise<Uint8Array>;
  destroy(): void;
}

let clientId = 81_300;

beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  );
  await preloadOpcWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')))
  );
}, 30_000);

const hostComments = (document: Document): Comment[] => document.package.document.comments ?? [];

async function editorSave(
  session: YrsSession,
  base: Document,
  comments: Comment[],
  host?: Document,
  stories: readonly string[] = []
) {
  const dirty = new DirtyProjectionStories();
  for (const story of ['body', ...stories]) dirty.add(story);
  const projected = yrsToDocument(session, host ? mergeDocxHostMetadata(base, host) : base, dirty.projectionOptions());
  return { projected, bytes: new Uint8Array(await saveEditorDocument(session, projected, comments)) };
}

async function open(route: Route, bytes: Uint8Array): Promise<Opened> {
  if (route === 'path that seeds in Rust') {
    const session = await createYrsSession({ clientId: (clientId += 1) });
    const host = session.openDocx(bytes.slice(), true).document;
    let projected: Document | undefined;
    return {
      session,
      lengths: (story) => session.paragraphSpans(story),
      revisions: () => session.listRevisions().length,
      save: async (comments, stories) => {
        const saved = await editorSave(session, projected ?? session.materializeDocx()!, comments(hostComments(host)), host, stories);
        projected = saved.projected;
        return saved.bytes;
      },
      destroy: () => session.destroy(),
    };
  }
  if (route === 'path that seeds in TypeScript') {
    const session = await createYrsSession({ clientId: (clientId += 1) });
    const parsed = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
    documentToYrs(session, parsed);
    const base = { ...parsed, originalBuffer: parsed.originalBuffer ?? bytes.slice().buffer };
    return {
      session,
      lengths: (story) => session.paragraphSpans(story),
      revisions: () => session.listRevisions().length,
      save: async (comments, stories) =>
        (await editorSave(session, base, comments(hostComments(parsed)), undefined, stories)).bytes,
      destroy: () => session.destroy(),
    };
  }
  const engine = await createResidentEngineSession(undefined, (clientId += 1));
  const hostJson = engine.openDocx(bytes.slice(), undefined, 'comment-range');
  const record: ResidentSaveRecord = { full: false };
  return {
    session: engine,
    lengths: (story) => engine.geometryReader.paragraphSpans(story),
    revisions: () => engine.proposalEngine.listRevisions().length,
    save: async (comments) => {
      const listed = comments(hostComments(decodeDocxHostJson(hostJson, bytes.slice()).document));
      return new Uint8Array(await engine.save(bytes.slice(), hostJson, undefined, listed, record));
    },
    destroy: () => engine.destroy(),
  };
}

function act({ session, lengths }: Opened, built: Built, action: ActionName): (list: Comment[]) => Comment[] {
  const typing = built.typing[action];
  if (typing) {
    const { paraId } = session.paragraphs(typing.story)[typing.paragraph]!;
    const end = () => lengths(typing.story).find((paragraph) => paragraph.paraId === paraId)!.length;
    const offset = typing.offset === 'end' ? end() : typing.offset;
    session.insertText({ story: typing.story, paraId, offset }, 'X');
  }
  if (action === 'bold where the range starts') {
    const { paraId } = session.paragraphs(built.bold.story)[0]!;
    session.toggleMark(
      { story: built.bold.story, start: { paraId, offset: 0 }, end: { paraId, offset: built.bold.to } },
      { type: 'bold' }
    );
  }
  const removed =
    action === 'deleting another comment' ? built.other : action === 'deleting the comment' ? built.ids[0] : undefined;
  if (removed === undefined) return (list) => list;
  session.applyRawOps('body', [{ op: 'removeComment', id: String(removed) }]);
  return (list) => list.filter((comment) => comment.id !== removed && comment.parentId !== removed);
}

const documentBody = (bytes: Uint8Array) => {
  const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
  return xml.slice(xml.indexOf('<w:body>'), xml.lastIndexOf('<w:sectPr'));
};

interface Marker {
  kind: 'Start' | 'End' | 'Reference';
  id: number;
  text: number;
  xml: number;
}

function scan(xml: string): { text: string; markers: Marker[] } {
  let text = '';
  const markers: Marker[] = [];
  const pattern = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:comment(RangeStart|RangeEnd|Reference)\b[^>]*?\bw:id="(\d+)"/g;
  for (const match of xml.matchAll(pattern)) {
    if (match[1] !== undefined) {
      text += match[1];
      continue;
    }
    const kind = match[2] === 'RangeStart' ? 'Start' : match[2] === 'RangeEnd' ? 'End' : 'Reference';
    markers.push({ kind, id: Number(match[3]), text: text.length, xml: match.index! });
  }
  return { text, markers };
}

const count = (xml: string, pattern: RegExp) => (xml.match(pattern) ?? []).length;

function container(xml: string, at: number): string {
  const before = xml.slice(0, at);
  const cells = count(before, /<w:tc[\s>]/g);
  if (cells > count(before, /<\/w:tc>/g)) return `cell ${cells}`;
  if (count(before, /<w:txbxContent>/g) > count(before, /<\/w:txbxContent>/g)) return 'text box';
  return 'body';
}

/** Where the comment's markers sit in the saved body, as a short summary. */
function placed(xml: string, id: number) {
  const { text, markers } = scan(xml);
  const own = markers.filter((marker) => marker.id === id);
  const first = (kind: Marker['kind']) => own.find((marker) => marker.kind === kind);
  const [opening, closing, reference] = [first('Start'), first('End'), first('Reference')];
  return {
    marks: own.map((marker) => marker.kind).join(' '),
    before: opening ? text.slice(Math.max(0, opening.text - 6), opening.text) : null,
    text: opening && closing ? text.slice(opening.text, closing.text) : null,
    after: closing ? text.slice(closing.text, closing.text + 6) : null,
    reference:
      closing && reference ? reference.xml > closing.xml && !xml.slice(closing.xml, reference.xml).includes('<w:t') : false,
    in: closing ? container(xml, closing.xml) : null,
  };
}

const marksOf = (xml: string, id: number) =>
  scan(xml)
    .markers.filter((marker) => marker.id === id)
    .map((marker) => marker.kind);
const paragraphsOf = (xml: string) => xml.match(/<w:p\b[^>]*>(?:(?!<w:p\b)[\s\S])*?<\/w:p>/g) ?? [];

/** The comment's marks, and where they sit in the paragraph that holds its start. */
function clamped(xml: string, id: number) {
  const paragraph = paragraphsOf(xml).find((candidate) => marksOf(candidate, id).includes('Start')) ?? '';
  const { text, markers } = scan(paragraph);
  const own = markers.filter((marker) => marker.id === id);
  const closing = own.find((marker) => marker.kind === 'End');
  const reference = own.find((marker) => marker.kind === 'Reference');
  return {
    marks: marksOf(xml, id).join(' '),
    inItsParagraph: own.map((marker) => marker.kind).join(' '),
    endsTheParagraph: closing ? closing.text === text.length : false,
    referenceFollows:
      closing && reference
        ? reference.xml > closing.xml && !paragraph.slice(closing.xml, reference.xml).includes('<w:t')
        : false,
  };
}

/** A range cut at the end of the paragraph it starts in, its reference right after. */
const CUT = { marks: 'Start End Reference', inItsParagraph: 'Start End Reference', endsTheParagraph: true, referenceFollows: true };

const stripIds = (xml: string) => xml.replace(/\s(?:w14:paraId|w14:textId)="[^"]*"/g, '');
const lastParagraph = (xml: string, text: string) =>
  stripIds((xml.match(/<w:p\b[^>]*>(?:(?!<w:p\b)[\s\S])*?<\/w:p>/g) ?? []).find((p) => p.includes(`>${text}<`)) ?? '');

const EDITS: ActionName[] = [
  'no edit',
  'typing where the range starts',
  'typing in the next paragraph',
  'typing where the range ends',
  'bold where the range starts',
  'deleting another comment',
];
const KEPT_ELSEWHERE: ActionName[] = ['no edit', 'typing in the next paragraph', 'bold where the range starts'];

const KEPT_FROM_SOURCE = (built: Built, route: Route, action: ActionName) =>
  built.ids.length > 1 && action === 'typing where the range ends' && KEEPS_BYTES(route);

const cells = BUILT.flatMap((built) =>
  EDITS.filter((action) => !action.startsWith('typing') || built.typing[action]).flatMap((action) =>
    ROUTES.filter((route) => !KEPT_FROM_SOURCE(built, route, action)).map(
      (route) => [built.name, route, action, built] as [string, Route, ActionName, Built]
    )
  )
);

describe('a comment range whose end lies in another story', () => {
  it.each(cells)(
    '%s, opened by the %s, after %s, keeps one start, one end and one reference in their places',
    async (_name, route, action, built) => {
      const source = documentBody(built.bytes);
      const opened = await open(route, built.bytes);
      try {
        const comments = act(opened, built, action);
        const xml = documentBody(await opened.save(comments));
        const typing = built.typing[action];
        for (const id of built.ids) {
          const want = { ...placed(source, id) };
          if (typing?.text) want.text = typing.text;
          expect([id, placed(xml, id)]).toEqual([id, want]);
        }
        if (KEEPS_BYTES(route) && KEPT_ELSEWHERE.includes(action)) {
          expect(lastParagraph(xml, built.last)).toBe(lastParagraph(source, built.last));
        }
      } finally {
        opened.destroy();
      }
    },
    TIMEOUT
  );

  it.each(BUILT.flatMap((built) => ROUTES.map((route) => [built.name, route, built] as [string, Route, Built])))(
    '%s, opened by the %s, loses every marker of the comment when it is deleted',
    async (_name, route, built) => {
      const source = documentBody(built.bytes);
      const opened = await open(route, built.bytes);
      try {
        const comments = act(opened, built, 'deleting the comment');
        const xml = documentBody(await opened.save(comments));
        const [removed, ...kept] = built.ids;
        expect(placed(xml, removed!).marks).toBe('');
        for (const id of [...kept, built.other]) {
          expect([id, placed(xml, id).marks]).toEqual([id, placed(source, id).marks]);
        }
      } finally {
        opened.destroy();
      }
    },
    TIMEOUT
  );
});

const anchorsOf = (session: Pick<YrsSession, 'resolveComment'>, id: number) => {
  try {
    return session
      .resolveComment(String(id))
      .slice()
      .sort((a, b) => a.story.localeCompare(b.story) || a.start - b.start);
  } catch {
    return [];
  }
};
const shownIn = (session: Pick<YrsSession, 'resolveComment'>, id: number) => [
  ...new Set(anchorsOf(session, id).map((anchor) => anchor.story)),
];
const listedIds = (session: Pick<YrsSession, 'listComments'>) =>
  session
    .listComments()
    .map((comment) => comment.id)
    .sort();

describe('a comment range end whose start is not in its story', () => {
  it.each(BUILT.flatMap((built) => ROUTES.map((route) => [built.name, route, built] as [string, Route, Built])))(
    '%s, opened by the %s, is shown and listed only as opened after an edit and a save',
    async (_name, route, built) => {
      const opened = await open(route, built.bytes);
      try {
        const action = built.typing['typing where the range ends'] ? 'typing where the range ends' : 'no edit';
        await opened.save(act(opened, built, action));
        expect([
          built.ids.map((id) => shownIn(opened.session, id)),
          listedIds(opened.session),
          opened.revisions(),
        ]).toEqual([built.ids.map((id) => built.anchors[id]), built.listed, 0]);
      } finally {
        opened.destroy();
      }
    },
    TIMEOUT
  );

  const LONE = (id: number, ids: number[]) =>
    docx(para(run('Plain one.')) + para(`${run('Lone')}${end(id)}${run(' end.')}`) + para(run('Last.')), ids);
  const ORPHANS: Array<[string, string, number, Uint8Array]> = [
    [
      'an end whose comment has no start anywhere',
      'is saved with no start, its end kept only where its paragraph keeps its source marks',
      3,
      LONE(3, [3]),
    ],
    [
      'an end whose id has no comment',
      'is saved with no start, its end kept only where its paragraph keeps its source marks',
      9,
      LONE(9, []),
    ],
  ];

  it.each(
    ORPHANS.flatMap(([name, outcome, id, bytes]) =>
      ROUTES.map((route) => [name, route, outcome, id, bytes] as [string, Route, string, number, Uint8Array])
    )
  )(
    '%s, opened by the %s, %s after an edit in its paragraph, and the session lists and shows nothing new for it',
    async (_name, route, _outcome, id, bytes) => {
      const opened = await open(route, bytes);
      try {
        const { paraId } = opened.session.paragraphs('body')[1]!;
        opened.session.insertText({ story: 'body', paraId, offset: 1 }, 'X');
        const xml = documentBody(await opened.save((list) => list));
        const want = KEEPS_BYTES(route) ? 'End Reference' : 'Reference';
        expect([
          marksOf(xml, id).join(' '),
          shownIn(opened.session, id),
          listedIds(opened.session).includes('9'),
          opened.revisions(),
        ]).toEqual([want, [], false, 0]);
      } finally {
        opened.destroy();
      }
    },
    TIMEOUT
  );
});

describe('a session joined from an update alone', () => {
  it(
    'holds no source document, so saving a comment that ends in a table cell after an edit fails as on main',
    async () => {
      const source = await createYrsSession({ clientId: (clientId += 1) });
      const peer = await createYrsSession({ clientId: (clientId += 1) });
      try {
        source.openDocx(BUILT[0]!.bytes.slice(), true);
        peer.loadState(source.encodeState());
        const { paraId } = peer.paragraphs('body')[0]!;
        peer.insertText({ story: 'body', paraId, offset: 7 }, 'X');
        let failure = '';
        try {
          await saveYrsDocx(peer, { updateModifiedDate: false });
        } catch (error) {
          failure = error instanceof Error ? error.message : String(error);
        }
        expect([peer.paragraphs('body')[0]!.text, peer.materializeDocx()?.originalBuffer, failure]).toEqual([
          'Alpha oXne.',
          undefined,
          'saveYrsDocx requires a session opened from DOCX bytes',
        ]);
      } finally {
        source.destroy();
        peer.destroy();
      }
    },
    TIMEOUT
  );
});

describe('the anchors both loaders give a comment range that crosses stories', () => {
  it.each(BUILT.map((built) => [built.name, built] as [string, Built]))(
    '%s gets the same anchors, in the story where it starts, from the Rust and the TypeScript loader',
    async (_name, built) => {
      const rust = await createYrsSession({ clientId: (clientId += 1) });
      const typescript = await createYrsSession({ clientId: (clientId += 1) });
      try {
        rust.openDocx(built.bytes.slice(), true);
        documentToYrs(typescript, await parseDocx(built.bytes.slice().buffer, { preloadFonts: false }));
        for (const id of built.ids) {
          expect([id, shownIn(rust, id)]).toEqual([id, built.anchors[id]]);
          expect([id, anchorsOf(typescript, id)]).toEqual([id, anchorsOf(rust, id)]);
        }
      } finally {
        rust.destroy();
        typescript.destroy();
      }
    },
    TIMEOUT
  );
});

const BEFORE_ROW = para(`${start(2)}${run('Before')}${end(2)}${run(' row.')}`);
const ACROSS_CELLS = docx(
  BEFORE_ROW + table(para(`${start(1)}${run('Left cell text.')}`), para(`${end(1)}${run('Right cell text.')}`)) +
    para(run('After row.')),
  [1, 2]
);
const BACKWARDS = docx(
  BEFORE_ROW + table(para(`${run('Left')}${end(1)}${run(' cell.')}`), para(`${start(1)}${run('Right cell text.')}`)) +
    para(run('After row.')),
  [1, 2]
);
const BETWEEN_PARAGRAPHS = docx(
  para(`${run('Alpha ')}${start(1)}${run('one.')}`) +
    '<w:commentRangeEnd w:id="1"/>' +
    para(`<w:r><w:commentReference w:id="1"/></w:r>${run('Beta two.')}`) +
    para(`${start(2)}${run('Delta')}${end(2)}${run(' four.')}`),
  [1, 2]
);
const FIELD =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> REF Mark \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:hyperlink w:anchor="Mark">${run('Link')}<w:r><w:commentReference w:id="1"/></w:r></w:hyperlink>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const IN_FIELD = docx(
  para(OPENING) +
    para(run('Beta two.')) +
    table(para(`${run('Gamma')}<w:commentRangeEnd w:id="1"/>${FIELD}${run(' three.')}`)) +
    para(run('Delta four.')),
  [1, 2]
);
const IN_HEADER = docx(
  para(`${run('Alpha ')}${start(7)}${run('seven')}${end(7)}${run('.')}`) + para(`${start(2)}${run('Beta')}${end(2)}${run(' two.')}`),
  [2, 7],
  para(`${run('Head')}${end(7)}${run(' text.')}`)
);
const IN_CONTROL = docx(
  para(OPENING) +
    para(run('Beta two.')) +
    `<w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent>${para(ENDING)}</w:sdtContent></w:sdt>` +
    para(run('Delta four.')),
  [1, 2]
);
const FROM_TEXT_BOX = docx(
  para(`${run('Alpha ')}${start(2)}${run('one')}${end(2)}${run('.')}`) +
    para(`${run('Box ')}${textBox(para(`${run('Inside ')}${start(1)}${run('box.')}`))}`) +
    para(ENDING) +
    para(run('Delta four.')),
  [1, 2]
);
const HEADER = 'hf:rIdH';

const OTHER = 2;
const withoutOther = (list: Comment[]) => list.filter((comment) => comment.id !== OTHER);
const removeOther = ({ session }: Opened) => session.applyRawOps('body', [{ op: 'removeComment', id: String(OTHER) }]);
const typeIn = ({ session }: Opened, story: string, paragraph: number, offset: number) => {
  const { paraId } = session.paragraphs(story)[paragraph]!;
  session.insertText({ story, paraId, offset }, 'X');
};
const cell = (column: number) => ({ story: 'body', tableIndex: 0, row: 0, column });
const replyTo = (list: Comment[], parentId: number): Comment => {
  const parent = list.find((comment) => comment.id === parentId)!;
  return { id: 10, author: parent.author, initials: parent.initials, date: parent.date, content: parent.content, parentId };
};
const headerXml = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/header1.xml']);
const counted = (xml: string, id: number) => {
  const kinds = marksOf(xml, id);
  const of = (kind: Marker['kind']) => kinds.filter((entry) => entry === kind).length;
  return { starts: of('Start'), ends: of('End'), references: of('Reference') };
};
const fromTextBox = (xml: string) => {
  const own = scan(xml).markers.filter((marker) => marker.id === 1);
  const at = (kind: Marker['kind']) => own.find((marker) => marker.kind === kind);
  const [opening, closing] = [at('Start'), at('End')];
  return [
    own.map((marker) => marker.kind).join(' '),
    opening ? container(xml, opening.xml) : null,
    closing ? container(xml, closing.xml) : null,
  ];
};

interface Crossing {
  name: string;
  outcome: string;
  bytes: Uint8Array;
  /** Edits and saves the opened document, as often as the case needs; the last save's bytes. */
  edit(opened: Opened): Promise<Uint8Array>;
  /** What the saved file shows, and what it should show. */
  check(saved: Uint8Array, route: Route): [unknown, unknown];
}

const IS_CUT = 'is saved cut at the end of the paragraph it starts in';
const IN_ORDER = 'is saved with one start, one end and one reference in order';
const cutIn = (saved: Uint8Array): [unknown, unknown] => [clamped(documentBody(saved), 1), CUT];

const CROSSINGS: Crossing[] = [
  {
    name: 'a comment into a table cell whose far paragraph is emptied',
    outcome: IS_CUT,
    bytes: BUILT[0]!.bytes,
    edit: async (opened) => {
      removeOther(opened);
      const { paraId } = opened.session.paragraphs(CELL_A)[0]!;
      const length = opened.lengths(CELL_A).find((paragraph) => paragraph.paraId === paraId)!.length;
      opened.session.deleteRange({ story: CELL_A, start: { paraId, offset: 0 }, end: { paraId, offset: length } });
      return opened.save(withoutOther);
    },
    check: cutIn,
  },
  {
    name: 'a comment from one cell to the next whose far column is deleted',
    outcome: IS_CUT,
    bytes: ACROSS_CELLS,
    edit: async (opened) => {
      removeOther(opened);
      opened.session.deleteColumn({ anchor: cell(1), head: cell(1) });
      return opened.save(withoutOther);
    },
    check: cutIn,
  },
  {
    name: 'a comment from one cell to the next whose two cells are merged',
    outcome: IS_CUT,
    bytes: ACROSS_CELLS,
    edit: async (opened) => {
      removeOther(opened);
      opened.session.mergeCells({ anchor: cell(0), head: cell(1) });
      return opened.save(withoutOther);
    },
    check: cutIn,
  },
  {
    name: 'a comment from one cell to the next, saved with its far column deleted, then undone, typed in and saved again',
    outcome: IN_ORDER,
    bytes: ACROSS_CELLS,
    edit: async (opened) => {
      opened.session.beginUndoCapture();
      removeOther(opened);
      opened.session.addUndoBoundary();
      opened.session.deleteColumn({ anchor: cell(1), head: cell(1) });
      await opened.save(withoutOther);
      opened.session.undo();
      typeIn(opened, CELL_A, 0, 2);
      return opened.save(withoutOther);
    },
    check: (saved) => [marksOf(documentBody(saved), 1).join(' '), 'Start End Reference'],
  },
  {
    name: 'a comment whose end sits directly in the body between two paragraphs',
    outcome: IS_CUT,
    bytes: BETWEEN_PARAGRAPHS,
    edit: async (opened) => {
      removeOther(opened);
      typeIn(opened, 'body', 0, 7);
      return opened.save(withoutOther);
    },
    check: cutIn,
  },
  {
    name: 'a comment into a table cell whose reference sits in a field result after its end',
    outcome: IS_CUT,
    bytes: IN_FIELD,
    edit: async (opened) => {
      removeOther(opened);
      typeIn(opened, CELL_A, 0, 2);
      return opened.save(withoutOther);
    },
    check: cutIn,
  },
  {
    name: 'a comment from one cell to the next whose start column is deleted',
    outcome: 'is saved without a range',
    bytes: ACROSS_CELLS,
    edit: async (opened) => {
      removeOther(opened);
      opened.session.deleteColumn({ anchor: cell(0), head: cell(0) });
      return opened.save(withoutOther);
    },
    check: (saved) => {
      const { starts, ends, references } = counted(documentBody(saved), 1);
      return [[starts, ends, references <= 1], [0, 0, true]];
    },
  },
  {
    name: 'a comment whose end and reference sit in the cell before its start',
    outcome: IS_CUT,
    bytes: BACKWARDS,
    edit: async (opened) => {
      removeOther(opened);
      typeIn(opened, CELL_B, 0, 2);
      return opened.save(withoutOther);
    },
    check: cutIn,
  },
  {
    name: 'a complete body comment whose id also ends in a header',
    outcome: 'keeps its range in the body and none in the header',
    bytes: IN_HEADER,
    edit: async (opened) => {
      removeOther(opened);
      typeIn(opened, HEADER, 0, 1);
      return opened.save(withoutOther, [HEADER]);
    },
    check: (saved) => [
      [marksOf(documentBody(saved), 7).join(' '), marksOf(headerXml(saved), 7).join(' ')],
      ['Start End Reference', ''],
    ],
  },
  {
    name: 'a comment into a block content control given a reply',
    outcome: 'is saved cut at the end of the paragraph it starts in, its reply with it',
    bytes: IN_CONTROL,
    edit: async (opened) => {
      removeOther(opened);
      typeIn(opened, 'body', 0, 7);
      return opened.save((list) => [...withoutOther(list), replyTo(list, 1)]);
    },
    check: (saved) => [[clamped(documentBody(saved), 1), clamped(documentBody(saved), 10)], [CUT, CUT]],
  },
  {
    name: 'a comment from a text box into a later paragraph, typed in where it ends',
    outcome: IN_ORDER,
    bytes: FROM_TEXT_BOX,
    edit: async (opened) => {
      removeOther(opened);
      typeIn(opened, 'body', 2, 2);
      return opened.save(withoutOther);
    },
    check: (saved) => [fromTextBox(documentBody(saved)), ['Start End Reference', 'text box', 'body']],
  },
  {
    name: 'a comment from a text box into a later paragraph, typed in the paragraph holding the text box',
    outcome: IN_ORDER,
    bytes: FROM_TEXT_BOX,
    edit: async (opened) => {
      removeOther(opened);
      typeIn(opened, 'body', 1, 1);
      return opened.save(withoutOther);
    },
    check: (saved) => [fromTextBox(documentBody(saved)), ['Start End Reference', 'text box', 'body']],
  },
  {
    name: 'a comment into a table cell whose reference sits in a field result, made bold where it starts',
    outcome: 'keeps the end in the cell beside the new one where the cell paragraph keeps its source bytes, as before',
    bytes: IN_FIELD,
    edit: async (opened) => {
      const { paraId } = opened.session.paragraphs('body')[0]!;
      opened.session.toggleMark({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 5 } }, { type: 'bold' });
      return opened.save((list) => list);
    },
    check: (saved, route) => [counted(documentBody(saved), 1).ends, KEEPS_BYTES(route) ? 2 : 1],
  },
];

describe('a comment range across stories saved after edits that move or remove its marks', () => {
  it.each(
    CROSSINGS.flatMap((crossing) =>
      ROUTES.map((route) => [crossing.name, route, crossing.outcome, crossing] as [string, Route, string, Crossing])
    )
  )(
    '%s, opened by the %s, %s',
    async (_name, route, _outcome, crossing) => {
      const opened = await open(route, crossing.bytes);
      try {
        const [received, wanted] = crossing.check(await crossing.edit(opened), route);
        expect(received).toEqual(wanted);
      } finally {
        opened.destroy();
      }
    },
    TIMEOUT
  );
});

describe('the source scan for comment ranges across stories', () => {
  it(
    'runs once for the opened document and not again for the projections change notifications make',
    async () => {
      const specifier = './crossStoryComments';
      const scanner: typeof import('./crossStoryComments') | null = await import(specifier).catch(() => null);
      expect(scanner === null ? 'no scanner' : 'a scanner').toBe('a scanner');
      if (!scanner) return;
      const scans = spyOn(scanner, 'crossStoryComments');
      const session = await createYrsSession({ clientId: (clientId += 1) });
      try {
        const host = session.openDocx(BUILT[0]!.bytes.slice(), true).document;
        let base = session.materializeDocx()!;
        const notify = () => {
          const dirty = new DirtyProjectionStories();
          dirty.add('body');
          base = yrsToDocument(session, mergeDocxHostMetadata(base, host), dirty.projectionOptions());
        };
        notify();
        for (const offset of [7, 8, 9]) {
          const { paraId } = session.paragraphs('body')[0]!;
          session.insertText({ story: 'body', paraId, offset }, 'X');
          notify();
        }
        await saveEditorDocument(session, base, hostComments(host));
        notify();
        expect(scans).toHaveBeenCalledTimes(1);
      } finally {
        scans.mockRestore();
        session.destroy();
      }
    },
    TIMEOUT
  );
});

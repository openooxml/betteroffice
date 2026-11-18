import { beforeAll, describe, expect, it } from 'bun:test';
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
import { createYrsSession, decodeDocxHostJson, type YrsSession } from './index';
import { createResidentEngineSession } from './residentEngineSession';
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

function docx(body: string, ids: number[]): Uint8Array {
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
        `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/></Relationships>`,
    ],
    [
      'word/styles.xml',
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
        '<w:name w:val="Normal"/></w:style></w:styles>',
    ],
    ['word/comments.xml', `<w:comments xmlns:w="${W}">${ids.map(comment).join('')}</w:comments>`],
    ['word/document.xml', `<w:document ${NAMESPACES}><w:body>${body}<w:sectPr/></w:body></w:document>`],
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
  offset: number;
  /** What the commented text, and the text right after the end, read once typed. */
  text?: string;
  after?: string;
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
  /** Each comment's anchors after opening, as `[story, open side]`. */
  anchors: Record<number, Array<[string, string | null]>>;
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
    anchors: { 1: [['body', 'end'], [CELL_A, 'start']] },
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
    anchors: { 1: [['body', null]] },
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
    anchors: { 1: [['body', 'end']] },
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
      'typing where the range ends': { story: CELL_B, paragraph: 0, offset: 2, after: 'RiXght' },
    },
    anchors: { 1: [[CELL_A, 'end'], [CELL_B, 'before']] },
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
      'typing where the range ends': { story: CELL_B, paragraph: 0, offset: 2, after: 'RiXght' },
    },
    anchors: {
      1: [[CELL_A, 'end'], [CELL_B, 'before']],
      4: [[CELL_A, 'end'], [CELL_B, 'start']],
      5: [[CELL_A, 'end'], [CELL_B, 'start']],
    },
  },
];

const ROUTES = ['path that seeds in Rust', 'worker engine', 'path that seeds in TypeScript'] as const;
type Route = (typeof ROUTES)[number];
const KEEPS_BYTES = (route: Route) => route !== 'path that seeds in TypeScript';

type Editable = Pick<YrsSession, 'paragraphs' | 'insertText' | 'toggleMark' | 'applyRawOps'>;

interface Opened {
  session: Editable;
  save(comments: (list: Comment[]) => Comment[]): Promise<Uint8Array>;
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

async function editorSave(session: YrsSession, base: Document, comments: Comment[], host?: Document) {
  const dirty = new DirtyProjectionStories();
  dirty.add('body');
  const projected = yrsToDocument(session, host ? mergeDocxHostMetadata(base, host) : base, dirty.projectionOptions());
  return new Uint8Array(await saveEditorDocument(session, projected, comments));
}

async function open(route: Route, bytes: Uint8Array): Promise<Opened> {
  if (route === 'path that seeds in Rust') {
    const session = await createYrsSession({ clientId: (clientId += 1) });
    const host = session.openDocx(bytes.slice(), true).document;
    return {
      session,
      save: async (comments) => editorSave(session, session.materializeDocx()!, comments(hostComments(host)), host),
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
      save: async (comments) => editorSave(session, base, comments(hostComments(parsed))),
      destroy: () => session.destroy(),
    };
  }
  const engine = await createResidentEngineSession(undefined, (clientId += 1));
  const hostJson = engine.openDocx(bytes.slice(), undefined, 'comment-range');
  return {
    session: engine,
    save: async (comments) => {
      const listed = comments(hostComments(decodeDocxHostJson(hostJson, bytes.slice()).document));
      return new Uint8Array(await engine.save(bytes.slice(), hostJson, undefined, listed, { full: false }));
    },
    destroy: () => engine.destroy(),
  };
}

function act(session: Editable, built: Built, action: ActionName): (list: Comment[]) => Comment[] {
  const typing = built.typing[action];
  if (typing) {
    const { paraId } = session.paragraphs(typing.story)[typing.paragraph]!;
    session.insertText({ story: typing.story, paraId, offset: typing.offset }, 'X');
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

const cells = BUILT.flatMap((built) =>
  EDITS.filter((action) => !action.startsWith('typing') || built.typing[action]).flatMap((action) =>
    ROUTES.map((route) => [built.name, route, action, built] as [string, Route, ActionName, Built])
  )
);

describe('a comment range whose end lies in another story', () => {
  it.each(cells)(
    '%s, opened by the %s, after %s, keeps one start, one end and one reference in their places',
    async (_name, route, action, built) => {
      const source = documentBody(built.bytes);
      const opened = await open(route, built.bytes);
      try {
        const comments = act(opened.session, built, action);
        const xml = documentBody(await opened.save(comments));
        const typing = built.typing[action];
        for (const id of built.ids) {
          const want = { ...placed(source, id) };
          if (typing?.text) want.text = typing.text;
          if (typing?.after) want.after = typing.after;
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
        const comments = act(opened.session, built, 'deleting the comment');
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

describe('the anchors both loaders give a comment range that crosses stories', () => {
  const anchorsOf = (session: YrsSession, id: number) =>
    session
      .resolveComment(String(id))
      .slice()
      .sort((a, b) => a.story.localeCompare(b.story) || a.start - b.start);

  it.each(BUILT.map((built) => [built.name, built] as [string, Built]))(
    '%s is seeded with the same anchors and open sides by the Rust and the TypeScript loader',
    async (_name, built) => {
      const rust = await createYrsSession({ clientId: (clientId += 1) });
      const typescript = await createYrsSession({ clientId: (clientId += 1) });
      try {
        rust.openDocx(built.bytes.slice(), true);
        documentToYrs(typescript, await parseDocx(built.bytes.slice().buffer, { preloadFonts: false }));
        for (const id of built.ids) {
          const seeded = anchorsOf(rust, id);
          expect([id, seeded.map((anchor) => [anchor.story, anchor.open ?? null])]).toEqual([id, built.anchors[id]]);
          expect([id, anchorsOf(typescript, id)]).toEqual([id, seeded]);
        }
      } finally {
        rust.destroy();
        typescript.destroy();
      }
    },
    TIMEOUT
  );
});

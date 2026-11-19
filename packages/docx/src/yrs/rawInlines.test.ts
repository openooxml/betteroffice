import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { saveEditorDocument, type EditorSaveRecord } from './editorSave';
import { createYrsSession, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

const run = (inner: string) => `<w:r>${inner}</w:r>`;
const preserved = (value: string) => (/^ | $| {2}/.test(value) ? ' xml:space="preserve"' : '');

type Spelling = { t: (value: string) => string };
const SOURCE: Spelling = { t: value => run(`<w:t xml:space="preserve">${value}</w:t>`) };
const REWRITTEN: Spelling = { t: value => run(`<w:t${preserved(value)}>${value}</w:t>`) };

const ZZZ = `<w:zzz>${run('<w:t>inner</w:t>')}</w:zzz>`;
const SUB_DOC = '<w:subDoc r:id="rId9"/>';
const INS_START = '<w:customXmlInsRangeStart w:id="7" w:author="Grace"/>';
const INS_END = '<w:customXmlInsRangeEnd w:id="7"/>';
const DEL_END = '<w:customXmlDelRangeEnd w:id="3"/>';
const MOVE_START = '<w:customXmlMoveFromRangeStart w:id="4" w:author="Grace"/>';

const paragraph = (inner: string) => `<w:p w14:paraId="1A2B3C01" w14:textId="1A2B3C01">${inner}</w:p>`;
const other = (spelling: Spelling) => `<w:p w14:paraId="1A2B3C02" w14:textId="1A2B3C02">${spelling.t('Other')}</w:p>`;
const documentXml = (inner: string, spelling: Spelling = SOURCE) =>
  `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body>${paragraph(inner)}${other(spelling)}${SECTION}</w:body></w:document>`;

function docx(inner: string): Uint8Array {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/subDocument" Target="part.docx" TargetMode="External"/></Relationships>`)],
    ['word/document.xml', toBytes(documentXml(inner))],
  ])));
}

const part = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
const body = (xml: string) => xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'));
const paragraphsOf = (xml: string) => body(xml).match(/<w:p\b(?:[^>]*\/>|[\s\S]*?<\/w:p>)/g) ?? [];

const TOKEN = /<w:(t|delText)\b[^>]*>([^<]*)<\/w:\1>|<w:(zzz|subDoc|customXml\w+Range(?:Start|End))\b(?:[^>]*\/>|[^>]*>([\s\S]*?)<\/w:\3>)/g;

/** A paragraph's text with each kept element in place: `[name]`, or `[zzz:text]` for the wrapper. */
function trace(paragraphXml: string): string {
  return [...paragraphXml.matchAll(TOKEN)].map(([, , text, kept, inner]) => {
    if (kept === undefined) return text;
    if (inner === undefined) return `[${kept}]`;
    return `[${kept}:${[...inner.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(([, value]) => value).join('')}]`;
  }).join('');
}

const traces = (xml: string) => paragraphsOf(xml).map(trace);

const editorRecords = new WeakMap<YrsSession, EditorSaveRecord>();

async function savedAsTheEditorSaves(session: YrsSession, base: Document): Promise<string> {
  const comments = base.package.document.comments ?? [];
  const record = editorRecords.get(session) ?? { full: false };
  editorRecords.set(session, record);
  const projected = yrsToDocument(session, base);
  return part(new Uint8Array(await saveEditorDocument(session, projected, comments, record)));
}

const sessions: YrsSession[] = [];
let clientId = 79100;

type Load = (bytes: Uint8Array) => Promise<{ session: YrsSession; base: Document }>;

const LOADERS: Array<[string, Load, Spelling, string]> = [
  ['the Rust loader', async bytes => {
    const session = await createYrsSession({ clientId: (clientId += 1) });
    sessions.push(session);
    session.openDocx(bytes, true);
    return { session, base: session.materializeDocx()! };
  }, SOURCE, 'the source document part unchanged'],
  ['the TypeScript loader', async bytes => {
    const session = await createYrsSession({ clientId: (clientId += 1) });
    sessions.push(session);
    const parsed = await parseDocx(bytes.slice().buffer as ArrayBuffer, { preloadFonts: false });
    documentToYrs(session, parsed);
    return { session, base: { ...parsed, originalBuffer: parsed.originalBuffer ?? (bytes.slice().buffer as ArrayBuffer) } };
  }, REWRITTEN, 'the document part as a rewritten paragraph spells it'],
];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { sessions.splice(0).forEach(session => session.destroy()); });

type Edit = (session: YrsSession, paraId: string) => void;

const typeAt = (offset: number): Edit => (session, paraId) => {
  session.insertText({ story: 'body', paraId, offset }, 'X');
};
const deleteBetween = (start: number, end: number): Edit => (session, paraId) => {
  session.deleteRange({ story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end } });
};
const splitAt = (offset: number): Edit => (session, paraId) => {
  session.splitParagraph({ story: 'body', paraId, offset });
};
const typeAndSplitAt = (offset: number): Edit => (session, paraId) => {
  typeAt(0)(session, paraId);
  splitAt(offset)(session, paraId);
};

const TYPESCRIPT_LOADER = 'the TypeScript loader';
const UNDO_LIMIT = 'limit: undo restores the mark at the start of the restored text on this loader, an end mark one character into it';

/**
 * An edit, the paragraphs it saves, for a split the paragraph that joining them back saves, and
 * where undoing the edit leaves the paragraph on the TypeScript loader when that is not the source.
 */
type Case = { name: string; edit: Edit; saved: string[]; joined?: string; undoneOnTheTypeScriptLoader?: string };

type Fixture = {
  name: string;
  inner: (spelling: Spelling) => string;
  text: string;
  embedded: boolean;
  loaders?: string[];
  cases: Case[];
};

const FIXTURES: Fixture[] = [
  {
    name: 'an unknown element holding a run',
    inner: s => `${s.t('Head ')}${ZZZ}${s.t(' tail')}`,
    text: 'Head  tail',
    embedded: true,
    cases: [
      { name: 'typing at the start of the paragraph', edit: typeAt(0), saved: ['XHead [zzz:inner] tail', 'Other'] },
      { name: 'deleting the first character', edit: deleteBetween(0, 1), saved: ['ead [zzz:inner] tail', 'Other'] },
      { name: 'deleting a range that covers the element', edit: deleteBetween(3, 8), saved: ['Heaail', 'Other'] },
      { name: 'deleting the range just before the element', edit: deleteBetween(2, 5), saved: ['He[zzz:inner] tail', 'Other'] },
      {
        name: 'splitting just before the element',
        edit: splitAt(5),
        saved: ['Head ', '[zzz:inner] tail', 'Other'],
        joined: 'Head [zzz:inner] tail',
      },
      {
        name: 'splitting just after the element',
        edit: splitAt(6),
        saved: ['Head [zzz:inner]', ' tail', 'Other'],
        joined: 'Head [zzz:inner] tail',
      },
    ],
  },
  {
    name: 'a w:subDoc',
    inner: s => `${s.t('Head ')}${SUB_DOC}${s.t(' tail')}`,
    text: 'Head  tail',
    embedded: false,
    cases: [
      { name: 'typing at the start of the paragraph', edit: typeAt(0), saved: ['XHead [subDoc] tail', 'Other'] },
      { name: 'deleting the first character', edit: deleteBetween(0, 1), saved: ['ead [subDoc] tail', 'Other'] },
      {
        name: 'deleting a range across the element',
        edit: deleteBetween(3, 7),
        saved: ['Hea[subDoc]ail', 'Other'],
        undoneOnTheTypeScriptLoader: 'Hea[subDoc]d  tail',
      },
      {
        name: 'splitting at the element',
        edit: splitAt(5),
        saved: ['Head ', '[subDoc] tail', 'Other'],
        joined: 'Head [subDoc] tail',
      },
    ],
  },
  {
    name: 'a customXmlInsRangeStart and End pair around Head',
    inner: s => `${INS_START}${s.t('Head')}${INS_END}${s.t(' tail')}`,
    text: 'Head tail',
    embedded: false,
    cases: [
      {
        name: 'typing at the start of the paragraph',
        edit: typeAt(0),
        saved: ['X[customXmlInsRangeStart]Head[customXmlInsRangeEnd] tail', 'Other'],
      },
      {
        name: 'deleting the first character',
        edit: deleteBetween(0, 1),
        saved: ['[customXmlInsRangeStart]ead[customXmlInsRangeEnd] tail', 'Other'],
      },
      {
        name: 'deleting a range across the end',
        edit: deleteBetween(2, 6),
        saved: ['[customXmlInsRangeStart]He[customXmlInsRangeEnd]ail', 'Other'],
        undoneOnTheTypeScriptLoader: '[customXmlInsRangeStart]Hea[customXmlInsRangeEnd]d tail',
      },
      {
        name: 'deleting everything between the start and the end',
        edit: deleteBetween(0, 4),
        saved: ['[customXmlInsRangeStart][customXmlInsRangeEnd] tail', 'Other'],
        undoneOnTheTypeScriptLoader: '[customXmlInsRangeStart]H[customXmlInsRangeEnd]ead tail',
      },
      {
        name: 'splitting at the start',
        edit: splitAt(0),
        saved: ['', '[customXmlInsRangeStart]Head[customXmlInsRangeEnd] tail', 'Other'],
        joined: '[customXmlInsRangeStart]Head[customXmlInsRangeEnd] tail',
      },
      {
        name: 'typing at the start and splitting at the end',
        edit: typeAndSplitAt(5),
        saved: ['X[customXmlInsRangeStart]Head[customXmlInsRangeEnd]', ' tail', 'Other'],
        joined: 'X[customXmlInsRangeStart]Head[customXmlInsRangeEnd] tail',
      },
    ],
  },
  {
    name: 'three adjacent empty elements',
    inner: s => `${s.t('Head ')}${DEL_END}${SUB_DOC}${MOVE_START}${s.t(' tail')}`,
    text: 'Head  tail',
    embedded: false,
    cases: [
      {
        name: 'typing at the start of the paragraph',
        edit: typeAt(0),
        saved: ['XHead [customXmlDelRangeEnd][subDoc][customXmlMoveFromRangeStart] tail', 'Other'],
      },
      {
        name: 'deleting the first character',
        edit: deleteBetween(0, 1),
        saved: ['ead [customXmlDelRangeEnd][subDoc][customXmlMoveFromRangeStart] tail', 'Other'],
      },
    ],
  },
  {
    name: 'a w:subDoc between two runs whose text seeds as one insert',
    inner: s => `${s.t('Hea')}${SUB_DOC}${s.t('d tail')}`,
    text: 'Head tail',
    embedded: false,
    loaders: ['the Rust loader'],
    cases: [
      { name: 'deleting a range across the element', edit: deleteBetween(2, 5), saved: ['He[subDoc]tail', 'Other'] },
    ],
  },
];

for (const fixture of FIXTURES) {
  describe(fixture.name, () => {
    const inner = fixture.inner(SOURCE);
    const source = trace(paragraph(inner));

    for (const [loader, load, spelling, saved] of LOADERS.filter(([name]) => fixture.loaders?.includes(name) ?? true)) {
      if (fixture.embedded) {
        it(`${loader}: the session seeds it as one embed whose text is neither shown nor found`, async () => {
          const { session } = await load(docx(inner));
          expect(session.paragraphs('body').map(entry => entry.text)).toEqual([fixture.text, 'Other']);
          expect(session.storySegments('body').filter(segment => segment.kind === 'embed')).toHaveLength(1);
          expect(session.paragraphSpans('body')[0]!.length).toBe(fixture.text.length + 1);
          expect(session.searchText('inner')).toEqual([]);
        });
      }

      for (const { name, edit, saved: edited, joined, undoneOnTheTypeScriptLoader } of fixture.cases) {
        const pinned = loader === TYPESCRIPT_LOADER ? undoneOnTheTypeScriptLoader : undefined;
        const undoing = pinned === undefined ? `undoing saves ${saved}` : `undoing moves it (${UNDO_LIMIT})`;
        it(`${loader}: ${name} keeps it in place, and ${undoing}`, async () => {
          const bytes = docx(inner);
          const unedited = await (async () => {
            const { session, base } = await load(bytes);
            return savedAsTheEditorSaves(session, base);
          })();
          expect(unedited).toBe(documentXml(fixture.inner(spelling), spelling));
          const { session, base } = await load(bytes);
          session.beginUndoCapture();
          const paraId = session.paragraphs('body')[0]!.paraId;
          edit(session, paraId);
          expect(traces(await savedAsTheEditorSaves(session, base))).toEqual(edited);
          if (joined !== undefined) {
            session.mergeParagraphs('body', paraId);
            expect(traces(await savedAsTheEditorSaves(session, base))).toEqual([joined, 'Other']);
          }
          expect(session.undo()).toBe(true);
          while (session.canUndo()) session.undo();
          if (pinned !== undefined) {
            expect(traces(await savedAsTheEditorSaves(session, base))).toEqual([pinned, 'Other']);
            return;
          }
          expect(traces(await savedAsTheEditorSaves(session, base))).toEqual([source, 'Other']);
          expect(await savedAsTheEditorSaves(session, base)).toBe(unedited);
        });
      }
    }
  });
}

const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
const pages = (xml: string) => xml.replaceAll('<w:br w:type="page"/>', '<w:t>[page]</w:t>');
const BREAK_RUN = /<w:r>(?:<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:br w:type="page"\/><\/w:r>/;

const BEFORE_A_BREAK: Array<[string, string]> = [['a w:subDoc', SUB_DOC], ['an unknown element holding a run', ZZZ]];

for (const [name, kept] of BEFORE_A_BREAK) {
  describe(`${name} before a page break in the same paragraph`, () => {
    const inner = `${SOURCE.t('Head ')}${kept}${SOURCE.t(' Mid')}${PAGE_BREAK}${SOURCE.t('End')}`;
    const keptTrace = trace(paragraph(kept));

    for (const [loader, load] of LOADERS) {
      it(`${loader}: typing before both keeps the break in its run and in place`, async () => {
        const { session, base } = await load(docx(inner));
        const first = session.paragraphs('body')[0]!;
        expect(first.properties._originalRunBoundaries).toHaveLength(4);
        typeAt(0)(session, first.paraId);
        const saved = paragraphsOf(await savedAsTheEditorSaves(session, base));
        expect(saved.map(entry => trace(pages(entry)))).toEqual([`XHead ${keptTrace} Mid[page]End`, 'Other']);
        expect(saved[0]).toMatch(BREAK_RUN);
      });
    }
  });
}

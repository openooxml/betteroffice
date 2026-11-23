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

const COMMENTS_XML =
  `<w:comments xmlns:w="${W}" xmlns:w14="${W14}"><w:comment w:id="1" w:author="Grace" w:date="2025-03-04T05:06:07Z" w:initials="G">` +
  '<w:p w14:paraId="0D000001" w14:textId="77777777"><w:r><w:t>Note 1</w:t></w:r></w:p></w:comment></w:comments>';

function docx(inner: string, comments = false): Uint8Array {
  const parts = new Map([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${comments ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : ''}</Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/subDocument" Target="part.docx" TargetMode="External"/>${comments ? `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/>` : ''}</Relationships>`)],
    ['word/document.xml', toBytes(documentXml(inner))],
  ]);
  if (comments) parts.set('word/comments.xml', toBytes(COMMENTS_XML));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
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

/** An edit, the paragraphs it saves, and for a split the paragraph that joining them back saves. */
type Case ={ name: string; edit: Edit; saved: string[]; joined?: string };

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
      { name: 'deleting a range across the element', edit: deleteBetween(3, 7), saved: ['Hea[subDoc]ail', 'Other'] },
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
      },
      {
        name: 'deleting everything between the start and the end',
        edit: deleteBetween(0, 4),
        saved: ['[customXmlInsRangeStart][customXmlInsRangeEnd] tail', 'Other'],
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

      for (const { name, edit, saved: edited, joined } of fixture.cases) {
        it(`${loader}: ${name} keeps it in place, and undoing saves ${saved}`, async () => {
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
          expect(traces(await savedAsTheEditorSaves(session, base))).toEqual([source, 'Other']);
          expect(await savedAsTheEditorSaves(session, base)).toBe(unedited);
        });
      }
    }
  });
}

describe('a w:subDoc after deleting across it, undoing, and typing elsewhere', () => {
  const inner = `${SOURCE.t('Head ')}${SUB_DOC}${SOURCE.t(' tail')}`;
  const [loader, load] = LOADERS.find(([name]) => name === 'the Rust loader')!;

  it(`${loader}: typing at the end saves the mark where it was before the deletion`, async () => {
    const { session, base } = await load(docx(inner));
    session.beginUndoCapture();
    const paraId = session.paragraphs('body')[0]!.paraId;
    deleteBetween(3, 7)(session, paraId);
    expect(traces(await savedAsTheEditorSaves(session, base))).toEqual(['Hea[subDoc]ail', 'Other']);
    expect(session.undo()).toBe(true);
    while (session.canUndo()) session.undo();
    expect(traces(await savedAsTheEditorSaves(session, base))).toEqual([trace(paragraph(inner)), 'Other']);
    typeAt(10)(session, paraId);
    expect(traces(await savedAsTheEditorSaves(session, base))).toEqual(['Head [subDoc] tailX', 'Other']);
  });
});

describe("a w:subDoc inside a comment range over 'middle'", () => {
  const inner =
    `${SOURCE.t('Head ')}<w:commentRangeStart w:id="1"/>${SOURCE.t('mid')}${SUB_DOC}${SOURCE.t('dle')}` +
    `<w:commentRangeEnd w:id="1"/>${run('<w:commentReference w:id="1"/>')}${SOURCE.t(' tail')}`;
  const RANGE = /<w:commentRangeStart\b[^>]*\bw:id="1"[^>]*\/>([\s\S]*?)<w:commentRangeEnd\b[^>]*\bw:id="1"[^>]*\/>/;

  const liveRange = (session: YrsSession) => {
    const [comment] = session.listComments();
    const [anchor] = session.resolveComment(comment!.id);
    let offset = 0;
    let text = '';
    for (const segment of session.storySegments(anchor!.story)) {
      if (segment.kind !== 'text') {
        offset += 1;
        continue;
      }
      text += segment.text.slice(Math.max(0, anchor!.start - offset), Math.max(0, anchor!.end - offset));
      offset += segment.text.length;
    }
    return text;
  };

  for (const [loader, load] of LOADERS) {
    it(`${loader}: deleting across the range start and the mark, then undoing, keeps the range on 'middle' and the mark in it`, async () => {
      const { session, base } = await load(docx(inner, true));
      const paraId = session.paragraphs('body')[0]!.paraId;
      expect(session.paragraphs('body')[0]!.text).toBe('Head middle tail');
      expect(liveRange(session)).toBe('middle');
      session.beginUndoCapture();
      deleteBetween(3, 10)(session, paraId);
      expect(session.paragraphs('body')[0]!.text).toBe('Heae tail');
      expect(session.undo()).toBe(true);
      while (session.canUndo()) session.undo();
      expect(session.paragraphs('body')[0]!.text).toBe('Head middle tail');
      expect(liveRange(session)).toBe('middle');
      const saved = await savedAsTheEditorSaves(session, base);
      expect(traces(saved)).toEqual(['Head mid[subDoc]dle tail', 'Other']);
      expect(trace(saved.match(RANGE)?.[1] ?? '')).toBe('mid[subDoc]dle');
    });
  }
});

const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
const pages = (xml: string) => xml.replaceAll('<w:br w:type="page"/>', '<w:t>[page]</w:t>');
const BREAK_RUN = /<w:r>(?:<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:br w:type="page"\/><\/w:r>/;

/** A kept element, and the recorded run boundaries of 'Head ', it, ' Mid', the break and 'End': its unit is one of them when it seeds as an embed. */
const BEFORE_A_BREAK: Array<[string, string, number]> = [
  ['a w:subDoc', SUB_DOC, 4],
  ['an unknown element holding a run', ZZZ, 5],
];

for (const [name, kept, boundaries] of BEFORE_A_BREAK) {
  describe(`${name} before a page break in the same paragraph`, () => {
    const inner = `${SOURCE.t('Head ')}${kept}${SOURCE.t(' Mid')}${PAGE_BREAK}${SOURCE.t('End')}`;
    const keptTrace = trace(paragraph(kept));

    for (const [loader, load] of LOADERS) {
      it(`${loader}: typing before both keeps the break in its run and in place`, async () => {
        const { session, base } = await load(docx(inner));
        const first = session.paragraphs('body')[0]!;
        expect(first.properties._originalRunBoundaries).toHaveLength(boundaries);
        typeAt(0)(session, first.paraId);
        const saved = paragraphsOf(await savedAsTheEditorSaves(session, base));
        expect(saved.map(entry => trace(pages(entry)))).toEqual([`XHead ${keptTrace} Mid[page]End`, 'Other']);
        expect(saved[0]).toMatch(BREAK_RUN);
      });

      it(`${loader}: typing at the end of the paragraph keeps the break in its run and in place`, async () => {
        const { session, base } = await load(docx(inner));
        const first = session.paragraphs('body')[0]!;
        typeAt(session.paragraphSpans('body')[0]!.length)(session, first.paraId);
        const saved = paragraphsOf(await savedAsTheEditorSaves(session, base));
        expect(saved.map(entry => trace(pages(entry)))).toEqual([`Head ${keptTrace} Mid[page]EndX`, 'Other']);
        expect(saved[0]).toMatch(BREAK_RUN);
      });
    }
  });
}

const FOREIGN = '<x:foo xmlns:x="urn:x"><x:bar/></x:foo>';
const texts = (xml: string) => [...xml.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map(([, value]) => value).join('');

describe("a foreign element holding content between 'Hel' and 'lo'", () => {
  const inner = `${SOURCE.t('Hel')}${FOREIGN}${SOURCE.t('lo')}`;

  for (const [loader, load] of LOADERS) {
    it(`${loader}: it takes no story unit and the paragraph is found as 'Hello'`, async () => {
      const { session } = await load(docx(inner));
      expect(session.paragraphs('body').map(entry => entry.text)).toEqual(['Hello', 'Other']);
      expect(session.paragraphSpans('body')[0]!.length).toBe(5);
      expect(session.storySegments('body').filter(segment => segment.kind === 'embed')).toEqual([]);
      expect(session.searchText('Hello').map(match => [match.start, match.end])).toEqual([[0, 5]]);
    });

    it(`${loader}: the structured export reports it as on main`, async () => {
      const { session } = await load(docx(inner));
      const read = session.exportStructured({ revisionView: 'accepted' });
      if (!read.ok) throw new Error(JSON.stringify(read.failure));
      const reported = loader === 'the Rust loader'
        ? ['An inline x:foo is not represented in the export; it stays in the package.']
        : [];
      expect(
        read.content.diagnostics.filter(({ code }) => code === 'unsupported-content').map(({ message }) => message),
      ).toEqual(reported);
      expect(JSON.stringify(read.content).split('"kind":"unsupported"')).toHaveLength(reported.length + 1);
    });

    it(`${loader}: typing at the start saves it at its original offset, as on main`, async () => {
      const { session, base } = await load(docx(inner));
      typeAt(0)(session, session.paragraphs('body')[0]!.paraId);
      const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
      expect(first!.split('<x:foo').map(texts)).toEqual(['XHe', 'llo']);
    });
  }
});

describe("an empty foreign element and a w:subDoc side by side between 'Hel' and 'lo'", () => {
  const inner = `${SOURCE.t('Hel')}<x:foo xmlns:x="urn:x"/>${SUB_DOC}${SOURCE.t('lo')}`;

  for (const [loader, load] of LOADERS) {
    it(`${loader}: typing at the start saves each once, the foreign one at its original offset and the w:subDoc at its anchor`, async () => {
      const { session, base } = await load(docx(inner));
      typeAt(0)(session, session.paragraphs('body')[0]!.paraId);
      const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
      expect(texts(first!)).toBe('XHello');
      expect([first!.split('<x:foo').length, first!.split('<w:subDoc').length]).toEqual([2, 2]);
      expect(first).toContain(SUB_DOC);
      expect(first!.split('<x:foo').map(texts)).toEqual(['XHe', 'llo']);
      expect(first!.split('<w:subDoc').map(texts)).toEqual(['XHel', 'lo']);
    });
  }
});

const split = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;
const SPLIT_FOREIGN: Array<[string, string]> = [
  ['an empty foreign element', '<x:foo xmlns:x="urn:x"/>'],
  ['a foreign element holding content', '<x:foo xmlns:x="urn:x"><x:bar/></x:foo>'],
];
const PARAGRAPH_OPEN = '<w:p w14:paraId="1A2B3C01" w14:textId="1A2B3C01">';

for (const [name, foreign] of SPLIT_FOREIGN) {
  describe(`${name} after the runs 'He' and 'l'`, () => {
    const inner = `${split('He')}${split('l')}${foreign}${split('lo')}`;

    for (const [loader, load] of LOADERS) {
      it(`${loader}: a paragraph alignment change saves the source runs as on main`, async () => {
        const { session, base } = await load(docx(inner));
        const paraId = session.paragraphs('body')[0]!.paraId;
        const position = { paraId, offset: 0 };
        session.setParagraphAttrs({ story: 'body', start: position, end: position }, { alignment: 'center' });
        const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
        expect(first).toBe(`${PARAGRAPH_OPEN}<w:pPr><w:jc w:val="center"/></w:pPr>${inner}</w:p>`);
      });

      it(`${loader}: typing a letter and deleting it saves the source runs as on main`, async () => {
        const { session, base } = await load(docx(inner));
        const paraId = session.paragraphs('body')[0]!.paraId;
        typeAt(0)(session, paraId);
        deleteBetween(0, 1)(session, paraId);
        const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
        expect(first).toBe(`${PARAGRAPH_OPEN}${inner}</w:p>`);
      });
    }
  });
}

const COLUMN_BREAK = '<w:r><w:br w:type="column"/></w:r>';
const flowBreaks = (xml: string) => pages(xml).replaceAll('<w:br w:type="column"/>', '<w:t>[column]</w:t>');
const flowBreakRun = (type: string) =>
  new RegExp(`<w:r>(?:<w:rPr>(?:(?!</w:rPr>)[\\s\\S])*</w:rPr>)?(?:<w:lastRenderedPageBreak/>)?<w:br w:type="${type}"/></w:r>`);
const spanLength = (session: YrsSession) => session.paragraphSpans('body')[0]!.length;

/** A break in a run of its own right after the kept element, and the loaders that seed it as a break. */
const LEADING_BREAKS: Array<[string, string, string[]]> = [
  ['page', PAGE_BREAK, ['the Rust loader', 'the TypeScript loader']],
  ['column', COLUMN_BREAK, ['the Rust loader']],
];

for (const [type, breakRun, loaders] of LEADING_BREAKS) {
  describe(`an unknown element holding a run directly before a ${type} break in a run of its own`, () => {
    const inner = `${ZZZ}${breakRun}${SOURCE.t('After')}`;
    const kept = trace(paragraph(ZZZ));
    const edits: Array<[string, Edit, string]> = [
      ['typing at the end', (session, paraId) => typeAt(spanLength(session))(session, paraId), `${kept}[${type}]AfterX`],
      ['changing the paragraph alignment', (session, paraId) => {
        const position = { paraId, offset: 0 };
        session.setParagraphAttrs({ story: 'body', start: position, end: position }, { alignment: 'center' });
      }, `${kept}[${type}]After`],
      ['deleting the element', (session, paraId) => {
        const at = spanLength(session) - 'After'.length - 1;
        deleteBetween(at, at + 1)(session, paraId);
      }, `[${type}]After`],
    ];

    for (const [loader, load] of LOADERS.filter(([name]) => loaders.includes(name))) {
      for (const [name, edit, expected] of edits) {
        it(`${loader}: ${name} keeps the break in its run and in place`, async () => {
          const { session, base } = await load(docx(inner));
          edit(session, session.paragraphs('body')[0]!.paraId);
          const saved = paragraphsOf(await savedAsTheEditorSaves(session, base));
          expect(saved.map(entry => trace(flowBreaks(entry)))).toEqual([expected, 'Other']);
          expect(saved[0]).toMatch(flowBreakRun(type));
        });
      }
    }
  });
}

const breakCounts = (xml: string) => ({
  page: xml.split('<w:br w:type="page"/>').length - 1,
  column: xml.split('<w:br w:type="column"/>').length - 1,
});
const centre: Edit = (session, paraId) => {
  const position = { paraId, offset: 0 };
  session.setParagraphAttrs({ story: 'body', start: position, end: position }, { alignment: 'center' });
};

/** A paragraph, the loaders that seed its breaks, an edit, and the breaks, text and markup its save holds. */
const LIVE_BREAKS: Array<{
  name: string;
  inner: string;
  loaders: string[];
  edit: Edit;
  breaks: { page: number; column: number };
  text: string;
  holds?: string[];
}> = [
  {
    name: 'typing before an unknown element that is directly before a leading page break keeps the page break and the element',
    inner: `${ZZZ}${PAGE_BREAK}${SOURCE.t('After')}`,
    loaders: ['the Rust loader', 'the TypeScript loader'],
    edit: typeAt(0),
    breaks: { page: 1, column: 0 },
    text: `X${trace(paragraph(ZZZ))}After`,
  },
  {
    name: 'typing before an unknown element that is directly before a leading column break keeps the column break and the element',
    inner: `${ZZZ}${COLUMN_BREAK}${SOURCE.t('After')}`,
    loaders: ['the Rust loader'],
    edit: typeAt(0),
    breaks: { page: 0, column: 1 },
    text: `X${trace(paragraph(ZZZ))}After`,
  },
  {
    name: 'typing at the end after a hyperlink holding only an unknown element before a leading page break keeps the page break; the element is lost, as on main',
    inner: `<w:hyperlink w:anchor="target">${ZZZ}</w:hyperlink>${PAGE_BREAK}${SOURCE.t('After')}`,
    loaders: ['the Rust loader', 'the TypeScript loader'],
    edit: (session, paraId) => typeAt(spanLength(session))(session, paraId),
    breaks: { page: 1, column: 0 },
    text: 'AfterX',
  },
  {
    name: 'changing the alignment of a paragraph of a page break, an unknown element and a column break keeps both breaks and the element',
    inner: `${PAGE_BREAK}${ZZZ}${COLUMN_BREAK}`,
    loaders: ['the Rust loader'],
    edit: centre,
    breaks: { page: 1, column: 1 },
    text: trace(paragraph(ZZZ)),
  },
  {
    name: 'deleting an unknown element before a leading page break followed by a tab keeps the page break and the tab',
    inner: `${ZZZ}${PAGE_BREAK}${run('<w:tab/>')}${SOURCE.t('After')}`,
    loaders: ['the Rust loader', 'the TypeScript loader'],
    edit: deleteBetween(0, 1),
    breaks: { page: 1, column: 0 },
    text: 'After',
    holds: ['<w:tab/>'],
  },
];

describe('a break the live paragraph holds beside an unknown element holding a run', () => {
  for (const { name, inner, loaders, edit, breaks, text, holds = [] } of LIVE_BREAKS) {
    for (const [loader, load] of LOADERS.filter(([entry]) => loaders.includes(entry))) {
      it(`${loader}: ${name}`, async () => {
        const { session, base } = await load(docx(inner));
        edit(session, session.paragraphs('body')[0]!.paraId);
        const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
        expect(breakCounts(first!)).toEqual(breaks);
        expect(trace(first!)).toBe(text);
        expect(holds.filter(markup => !first!.includes(markup))).toEqual([]);
      });
    }
  }
});

describe('runs of their own languages around an unknown element holding a run', () => {
  const lang = (value: string, content: string) => run(`<w:rPr><w:lang w:val="${value}"/></w:rPr>${content}`);
  const rule = '<w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
    '<v:rect style="width:0pt;height:1.5pt" o:hr="t" o:hrstd="t"/></w:pict>';
  const inner = `${lang('fr-FR', '<w:t>A</w:t>')}${ZZZ}${lang('de-DE', rule)}${lang('es-ES', '<w:t>B</w:t>')}`;
  const runWith = (xml: string, needle: string) => {
    const at = xml.indexOf(needle);
    return at < 0 ? '' : xml.slice(xml.lastIndexOf('<w:r>', at), xml.indexOf('</w:r>', at));
  };

  for (const [loader, load] of LOADERS) {
    it(`${loader}: typing at the end keeps each run's own language, the horizontal rule's too`, async () => {
      const { session, base } = await load(docx(inner));
      typeAt(spanLength(session))(session, session.paragraphs('body')[0]!.paraId);
      const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
      expect(trace(first!)).toBe(`A${trace(paragraph(ZZZ))}BX`);
      const languages = [['>A<', 'fr-FR'], ['<w:pict', 'de-DE'], ['>B', 'es-ES']]
        .map(([needle, value]) => [needle, runWith(first!, needle!).includes(`w:val="${value}"`)]);
      expect(languages).toEqual([['>A<', true], ['<w:pict', true], ['>B', true]]);
    });
  }
});

describe('a w:subDoc before a right-to-left scope', () => {
  const inner = `${SUB_DOC}<w:dir w:val="rtl">${SOURCE.t('مرحبا')}</w:dir>`;
  const INSIDE_A_DIR = /<w:dir\b[^>]*>(?:(?!<\/w:dir>)[\s\S])*<w:subDoc\b/;

  for (const [loader, load] of LOADERS) {
    it(`${loader}: typing before the scope saves the w:subDoc once and outside every w:dir`, async () => {
      const { session, base } = await load(docx(inner));
      typeAt(0)(session, session.paragraphs('body')[0]!.paraId);
      const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
      expect(texts(first!)).toBe('Xمرحبا');
      expect(first!.split('<w:subDoc').length).toBe(2);
      expect(first).not.toMatch(INSIDE_A_DIR);
    });
  }
});

describe('a leading page break without a kept element saves as on main', () => {
  const countPages = (xml: string) => xml.split('<w:br w:type="page"/>').length - 1;

  for (const [loader, load] of LOADERS) {
    it(`${loader}: splitting at the start of a paragraph that opens with a page break saves one page break, in the second half`, async () => {
      const { session, base } = await load(docx(`${PAGE_BREAK}${SOURCE.t('After')}`));
      session.splitParagraph({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 });
      const saved = paragraphsOf(await savedAsTheEditorSaves(session, base));
      expect(saved.map(countPages)).toEqual([0, 1, 0]);
      expect(texts(saved[1]!)).toBe('After');
    });

    it(`${loader}: typing after a hyperlink that opens with a page break saves the break as main does`, async () => {
      const { session, base } = await load(docx(`<w:hyperlink w:anchor="target">${PAGE_BREAK}${SOURCE.t('link')}</w:hyperlink>${SOURCE.t(' After')}`));
      typeAt(session.paragraphSpans('body')[0]!.length)(session, session.paragraphs('body')[0]!.paraId);
      const [first] = paragraphsOf(await savedAsTheEditorSaves(session, base));
      expect(texts(first!)).toBe('link AfterX');
      if (loader === 'the Rust loader') {
        expect(first).toMatch(/<w:hyperlink\b[^>]*><w:r><w:br w:type="page"\/><\/w:r>/);
      } else {
        expect(countPages(first!)).toBe(0);
      }
    });
  }
});

describe('kept empty elements spread over many paragraphs', () => {
  const texts = ['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon'];
  const placed = (index: number) => {
    const text = texts[index % texts.length]!;
    const at = index % (text.length + 1);
    return `${text.slice(0, at)}[subDoc]${text.slice(at)}`;
  };
  const source = (index: number) => {
    const text = texts[index % texts.length]!;
    const at = index % (text.length + 1);
    return `${SOURCE.t(text.slice(0, at))}${SUB_DOC}${SOURCE.t(text.slice(at))}`;
  };
  let inner = `${source(0)}${SUB_DOC}`;
  for (let index = 1; index < 40; index += 1) {
    const id = (0x20000000 + index).toString(16).toUpperCase();
    inner += `</w:p><w:p w14:paraId="${id}" w14:textId="${id}">${source(index)}`;
  }

  for (const [loader, load] of LOADERS) {
    it(`${loader}: typing in one paragraph saves every mark of every paragraph in its place`, async () => {
      const { session, base } = await load(docx(inner));
      typeAt(0)(session, session.paragraphs('body')[20]!.paraId);
      const saved = traces(await savedAsTheEditorSaves(session, base));
      const expected = Array.from({ length: 40 }, (_, index) => placed(index));
      expected[0] = `${expected[0]}[subDoc]`;
      expected[20] = `X${expected[20]}`;
      expect(saved).toEqual([...expected, 'Other']);
    });
  }
});
